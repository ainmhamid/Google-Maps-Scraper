# -*- coding: utf-8 -*-
import datetime
from playwright.sync_api import sync_playwright
from dataclasses import dataclass, asdict, field
import pandas as pd
import argparse
import os
import sys
import re
import html
import urllib.request
 
def sanitize_filename(name):
    """Sanitize business name to be safe for file system."""
    if not name:
        return "unnamed_business"
    # Remove invalid characters for Windows filenames
    sanitized = re.sub(r'[\\/*?:"<>|]', "", name)
    # Replace spaces and multiple underscores
    sanitized = sanitized.strip().replace(" ", "_")
    sanitized = re.sub(r'_+', '_', sanitized)
    return sanitized

def download_image(url, save_dir, filename):
    """Download image from url and save it to save_dir/filename."""
    if not url:
        return ""
    try:
        os.makedirs(save_dir, exist_ok=True)
        # Google user content image urls are sometimes parameters-heavy, let's get a stable suffix
        # If the user content image suffix is present, replace it with =s800 to get a reasonably large image
        if "googleusercontent.com" in url or "ggpht.com" in url:
            base_url = url.split('=')[0]
            url = f"{base_url}=s800"
            
        file_path = os.path.join(save_dir, filename)
        
        # Define headers to mimic browser request to avoid getting blocked
        req = urllib.request.Request(
            url, 
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'}
        )
        with urllib.request.urlopen(req, timeout=10) as response, open(file_path, 'wb') as out_file:
            out_file.write(response.read())
            
        # Return path relative to save_at directory (e.g. images_coffee_shops/filename.jpg)
        return os.path.join(os.path.basename(save_dir), filename)
    except Exception as e:
        print(f"Error downloading image {url}: {e}")
        return ""

@dataclass
class Business:
    """holds business data"""
    name: str = None
    type: str = None 
    price: str = None 
    address: str = None
    website: str = None
    phone_number: str = None
    category: str = None
    location: str = None
    reviews_count: int = None
    reviews_average: float = None
    image_url: str = None
    local_image_path: str = None
    
    review_1_username: str = None
    review_1_local_guide: str = None
    review_1_text: str = None
    review_1_images: str = None
    review_1_local_images: str = None
    
    review_2_username: str = None
    review_2_local_guide: str = None
    review_2_text: str = None
    review_2_images: str = None
    review_2_local_images: str = None
    
    review_3_username: str = None
    review_3_local_guide: str = None
    review_3_text: str = None
    review_3_images: str = None
    review_3_local_images: str = None
    
    def __hash__(self):
        """Make Business hashable for duplicate detection.
        Consider businesses different if:
        - Name is different, OR
        - Same name but different non-empty contact info (website/phone)
        """
        # Create a tuple of fields that must match for duplicates
        hash_fields = [self.name]
        # Only include contact info fields if they're not empty
        if getattr(self, "website", None):
            hash_fields.append(f"website:{self.website}")
        if getattr(self, "phone_number", None):
            hash_fields.append(f"phone:{self.phone_number}")

        return hash(tuple(hash_fields))

@dataclass
class BusinessList:
    """holds list of Business objects,
    and save to both excel and csv
    """
    business_list: list[Business] = field(default_factory=list)
    _seen_businesses: set = field(default_factory=set, init=False)
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    save_at = os.path.join('GMaps Data', today)
    os.makedirs(save_at, exist_ok=True)

    def add_business(self, business: Business):
        """Add a business to the list if it's not a duplicate based on key attributes"""
        business_hash = hash(business)
        if business_hash not in self._seen_businesses:
            self.business_list.append(business)
            self._seen_businesses.add(business_hash)
    
    def dataframe(self):
        """transform business_list to pandas dataframe

        Returns: pandas dataframe
        """
        df = pd.json_normalize(
            (asdict(business) for business in self.business_list), sep="_"
        )
        # Ensure all object columns are treated as string type to preserve Unicode
        for col in df.select_dtypes(include='object').columns:
            df[col] = df[col].astype(str)
        return df

    def save_to_excel(self, filename):
        """saves pandas dataframe to excel (xlsx) file

        Args:
            filename (str): filename
        """
        self.dataframe().to_excel(f"{self.save_at}/{filename}.xlsx", index=False, engine='openpyxl')

    def save_to_csv(self, filename):
        """saves pandas dataframe to csv file

        Args:
            filename (str): filename
        """
        self.dataframe().to_csv(f"{self.save_at}/{filename}.csv", index=False, encoding='utf-8')

def main():
    # Reconfigure standard output to support UTF-8 encoding (prevents UnicodeEncodeError on some Windows consoles)
    if sys.stdout.encoding != 'utf-8':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except AttributeError:
            pass

    # read search from arguments
    parser = argparse.ArgumentParser()
    parser.add_argument("-s", "--search", type=str)
    parser.add_argument("-t", "--total", type=int)
    args = parser.parse_args()
    
    if args.search:
        search_list = [args.search]
        
    if args.total:
        total = args.total
    else:
        # if no total is passed, we set the value to random big number
        total = 1_000_000

    if not args.search:
        search_list = []
        # read search from input.txt file
        input_file_name = 'input.txt'
        # Get the absolute path of the file in the current working directory
        input_file_path = os.path.join(os.getcwd(), input_file_name)
        # Check if the file exists
        if os.path.exists(input_file_path):
        # Open the file in read mode with UTF-8 encoding
            with open(input_file_path, 'r', encoding='utf-8') as file:
            # Read all lines into a list
                search_list = file.readlines()
                
        if len(search_list) == 0:
            print('Error occured: You must either pass the -s search argument, or add searches to input.txt')
            sys.exit()
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page(locale="en-GB")

        page.goto("https://www.google.com/maps", timeout=20000)
        
        # Handle consent banner if it appears
        page.wait_for_timeout(3000)
        try:
            consent_btn = page.locator('button:has-text("Accept all"), button:has-text("Agree"), button[aria-label*="Accept"]')
            if consent_btn.count() > 0:
                consent_btn.first.click()
                page.wait_for_timeout(3000)
        except Exception:
            pass
        
        for search_for_index, search_for in enumerate(search_list):
            search_for = search_for.strip()  # Remove newlines and whitespace
            print(f"-----\n{search_for_index} - {search_for}")

            # Robust searchbox locator with fallbacks
            search_box = page.locator('input[id="searchboxinput"], input[name="q"], input.searchboxinput')
            search_box.first.fill(search_for)
            page.wait_for_timeout(3000)

            page.keyboard.press("Enter")
            page.wait_for_timeout(5000)

            # scrolling
            page.hover('//a[contains(@href, "https://www.google.com/maps/place")]')

            previously_counted = 0
            while True:
                page.mouse.wheel(0, 10000)
                page.wait_for_timeout(3000)

                if (
                    page.locator(
                        '//a[contains(@href, "https://www.google.com/maps/place")]'
                    ).count()
                    >= total
                ):
                    listings = page.locator(
                        '//a[contains(@href, "https://www.google.com/maps/place")]'
                    ).all()[:total]
                    listings = [listing.locator("xpath=..") for listing in listings]
                    print(f"Total Scraped: {len(listings)}")
                    break
                else:
                    if (
                        page.locator(
                            '//a[contains(@href, "https://www.google.com/maps/place")]'
                        ).count()
                        == previously_counted
                    ):
                        listings = page.locator(
                            '//a[contains(@href, "https://www.google.com/maps/place")]'
                        ).all()
                        print(f"Arrived at all available\nTotal Scraped: {len(listings)}")
                        break
                    else:
                        previously_counted = page.locator(
                            '//a[contains(@href, "https://www.google.com/maps/place")]'
                        ).count()
                        print(
                            f"Currently Scraped: ",
                            page.locator(
                                '//a[contains(@href, "https://www.google.com/maps/place")]'
                            ).count(), end='\r'
                        )

            business_list = BusinessList()

            # scraping
            for listing in listings:
                try:                        
                    listing.click()
                    page.wait_for_timeout(2000)

                    name_attribute = 'h1.DUwDvf'
                    address_xpath = '//button[@data-item-id="address"]//div[contains(@class, "fontBodyMedium")]'
                    website_xpath = '//a[@data-item-id="authority"]//div[contains(@class, "fontBodyMedium")]'
                    phone_number_xpath = '//button[contains(@data-item-id, "phone:tel:")]//div[contains(@class, "fontBodyMedium")]'
                    type_xpath = '//button[contains(@jsaction, "category")]'
                    price_xpath = '//span[contains(normalize-space(.), "RM")]'
                    review_count_xpath = '//div[@jsaction="pane.reviewChart.moreReviews"]//span'
                    reviews_average_xpath = '//div[@jsaction="pane.reviewChart.moreReviews"]//div[@role="img"]' # or .fontDisplayLarge locator
                                                           
                    business = Business()
                   
                    if name_value := page.locator(name_attribute).inner_text():
                        # Decode HTML entities to support other languages
                        business.name = html.unescape(name_value).strip()
                    else:
                        business.name = ""

                    if page.locator(address_xpath).count() > 0:
                        business.address = page.locator(address_xpath).all()[0].inner_text()
                    else:
                        business.address = ""

                    business.website = page.url

                    if page.locator(phone_number_xpath).count() > 0:
                        business.phone_number = page.locator(phone_number_xpath).all()[0].inner_text()
                    else:
                        business.phone_number = ""

                    if page.locator(type_xpath).count() > 0:
                        business.type = page.locator(type_xpath).all()[0].inner_text()
                    else:
                        business.type = ""

                    if page.locator(price_xpath).count() > 0:
                        price_text = page.locator(price_xpath).all()[0].inner_text()
                        # Clean encoding artifacts: decode HTML entities, remove non-ASCII chars
                        price_text = html.unescape(price_text)
                        # Remove all non-printable/special characters except digits, RM, and dashes
                        price_text = ''.join(c for c in price_text if c.isascii() or c.isdigit() or c in 'RM- ')
                        # Extract just "RM" followed by price range (e.g., "RM 20–40")
                        price_match = re.search(r'RM\s*\d+\s*[-–]\s*\d+', price_text)
                        business.price = price_match.group(0).strip() if price_match else price_text.strip()
                    else:
                        business.price = ""
                        
                    if page.locator(review_count_xpath).count() > 0:
                        business.reviews_count = int(page.locator(review_count_xpath).inner_text().split()[0].replace(',', '').strip())
                    else:
                        business.reviews_count = ""
                        
                    if page.locator(reviews_average_xpath).count() > 0:
                        business.reviews_average = float(page.locator(reviews_average_xpath).get_attribute('aria-label').split()[0].replace(',', '.').strip())
                    else:
                        business.reviews_average = ""
                
                    business.category = search_for.split(' in ')[0].strip()
                    business.location = search_for.split(' in ')[-1].strip()

                    # COVER IMAGE EXTRACTION
                    cover_image_url = ""
                    try:
                        # Try button with jsaction matching heroHeaderImage
                        cover_locator = page.locator('button[jsaction*="heroHeaderImage"] img')
                        if cover_locator.count() > 0:
                            cover_image_url = cover_locator.first.get_attribute('src')
                        else:
                            # Fallback 1: button with jsaction matching mediaViewer
                            cover_locator = page.locator('button[jsaction*="mediaViewer"] img, button[jsaction*="media.viewer"] img')
                            if cover_locator.count() > 0:
                                cover_image_url = cover_locator.first.get_attribute('src')
                            else:
                                # Fallback 2: first image starting with googleusercontent/googleapis
                                img_locs = page.locator('img').all()
                                for img in img_locs:
                                    src = img.get_attribute('src')
                                    if src and ('googleusercontent' in src or 'googleapis.com/cbk' in src):
                                        if '=s32' not in src and '=s36' not in src and '=s40' not in src and '-c-k-no' not in src:
                                            cover_image_url = src
                                            break
                    except Exception as img_err:
                        print(f"Error extracting cover image: {img_err}")
                    
                    business.image_url = cover_image_url
                    
                    # COVER IMAGE DOWNLOAD
                    if cover_image_url:
                        images_dir_name = f"images_{search_for.replace(' ', '_')}"
                        save_images_path = os.path.join(business_list.save_at, images_dir_name)
                        sanitized_name = sanitize_filename(business.name)
                        image_filename = f"{sanitized_name}_cover.jpg"
                        
                        local_path = download_image(cover_image_url, save_images_path, image_filename)
                        business.local_image_path = local_path
                    else:
                        business.local_image_path = ""

                    # REVIEWS EXTRACTION
                    try:
                        print(f"Scraping reviews for {business.name}...")
                        # Auto-wait up to 5 seconds and click reviews tab
                        reviews_tab = page.locator('button[role="tab"]:has-text("Reviews")').first
                        reviews_tab.click(timeout=5000)
                        page.wait_for_timeout(2000)
                        
                        # Scroll reviews container to trigger loading
                        try:
                            scrollable_div = page.locator('div.m6QErb.dS8AEf:not(.ecceSd)').first
                            for _ in range(3):
                                scrollable_div.evaluate('(el) => el.scrollTop = el.scrollHeight')
                                page.wait_for_timeout(1000)
                        except Exception as se:
                            print(f"Scroll error: {se}")
                            page.wait_for_timeout(3000)
                        
                        # Wait for at least one review card to load
                        try:
                            page.wait_for_selector('.jftiEf', timeout=5000)
                        except Exception as we:
                            print(f"Wait for reviews timed out: {we}")
                        
                        # Find the first 3 review cards
                        review_cards = page.locator('.jftiEf').all()[:3]
                        print(f"Found {len(review_cards)} reviews.")
                        
                        for idx, card in enumerate(review_cards):
                            review_num = idx + 1
                            username = ""
                            local_guide = "No"
                            review_text = ""
                            review_imgs = []
                            local_review_imgs = []
                            
                            try:
                                # Username
                                user_loc = card.locator('.d4r55')
                                if user_loc.count() > 0:
                                    username = user_loc.first.inner_text().strip()
                                
                                # Local Guide status
                                card_text = card.inner_text()
                                local_guide = "Yes" if "Local Guide" in card_text else "No"
                                
                                # Review Text
                                text_loc = card.locator('.wiI7pd')
                                if text_loc.count() > 0:
                                    review_text = text_loc.first.inner_text().strip()
                                
                                # Review Images (extracted from computed style of button.Tya61d)
                                raw_imgs = card.evaluate("""(el) => {
                                    let urls = [];
                                    let buttons = el.querySelectorAll('button.Tya61d');
                                    for (let btn of buttons) {
                                        let bg = window.getComputedStyle(btn).backgroundImage;
                                        if (bg && bg !== 'none') {
                                            let match = bg.match(/url\\(["']?(.*?)["']?\\)/);
                                            if (match && match[1]) {
                                                urls.push(match[1]);
                                            }
                                        }
                                    }
                                    return urls;
                                }""")
                                review_imgs = raw_imgs[:3]
                                            
                                # Download review images
                                if review_imgs:
                                    images_dir_name = f"images_{search_for.replace(' ', '_')}"
                                    save_images_path = os.path.join(business_list.save_at, images_dir_name)
                                    sanitized_name = sanitize_filename(business.name)
                                    
                                    for img_idx, img_url in enumerate(review_imgs):
                                        img_filename = f"{sanitized_name}_review_{review_num}_photo_{img_idx+1}.jpg"
                                        loc_img_path = download_image(img_url, save_images_path, img_filename)
                                        if loc_img_path:
                                            local_review_imgs.append(loc_img_path)
                                            
                            except Exception as re_err:
                                print(f"Error scraping review {review_num}: {re_err}")
                                
                            setattr(business, f"review_{review_num}_username", username)
                            setattr(business, f"review_{review_num}_local_guide", local_guide)
                            setattr(business, f"review_{review_num}_text", review_text)
                            setattr(business, f"review_{review_num}_images", ",".join(review_imgs))
                            setattr(business, f"review_{review_num}_local_images", ",".join(local_review_imgs))
                            
                        # Click back to Overview tab
                        try:
                            overview_tab = page.locator('button[role="tab"]:has-text("Overview")').first
                            overview_tab.click(timeout=3000)
                            page.wait_for_timeout(1000)
                        except Exception:
                            pass
                    except Exception as reviews_err:
                        print(f"Error extracting reviews or tab: {reviews_err}")

                    business_list.add_business(business)
                except Exception as e:
                    print(f'Error occurred: {e}', end='\r')
            
            # output
            business_list.save_to_excel(f"{search_for}".replace(' ', '_'))
            business_list.save_to_csv(f"{search_for}".replace(' ', '_'))

        browser.close()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f'Failed err: {e}')