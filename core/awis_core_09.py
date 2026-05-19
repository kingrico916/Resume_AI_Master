"""
AWIS Core - Base scraper class
Extracted from GIT.PY production code
Contains all shared functionality and generic selector logic
"""
import time
import os
import re
import csv
import json
import logging
from datetime import datetime
from urllib.parse import urlparse, urljoin
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import openai
import anthropic

from selector_library import SelectorLibrary

# Custom Exceptions
class QuotaExceededError(Exception):
    """Raised when API quota is exceeded"""
    pass

class ProviderUnavailableError(Exception):
    """Raised when AI provider is unavailable"""
    pass

class AWISCore:
    """Base scraper with generic selector magic"""
    
    def __init__(self, output_dir="automation_output", args=None):
        self.args = args or type('obj', (object,), {
            'pipeline': 'api',
            'ai': 'chatgpt',
            'backup_policy': 'onfail',
            'verbose': False,
            'operator_email': 'anthony.antonucci@workforwarriors.org',
            'operator_name': 'ANTONUCCI'
        })()
        
        self.output_dir = output_dir
        self.setup_logging()
        
        self.csv_dir = os.path.join(output_dir, "csv_data")
        self.text_backup_dir = os.path.join(output_dir, "text_backup")
        self.ensure_directories()
        
        self.driver = None
        self.job_counter = 0
        self.successful_captures = 0
        self.failed_captures = 0
        self.api_failures = 0
        self.text_backups_saved = 0
        self.processed_req_ids = self.load_processed_ids()
        
        self.session_start_time = datetime.now()
        self.javascript_mode = False  # Adaptive extraction flag
        
        # AI clients
        self.openai_client = None
        self.claude_client = None
        self.init_ai_clients()
        
        self.csv_headers = [
            "Job Title", "Job Description", "Job Type", "Categories",
            "Location", "City", "State", "Country", "Zip Code", "Address",
            "Salary From", "Salary To", "Salary Period", "Apply Email",
            "Posting Date", "Expiration Date", "Qualifications",
            "Employer Email", "Full Name", "Company Name", "Employer Website",
            "Employer Phone"
        ]
        
        self.csv_filename = None
        
    def setup_logging(self):
        """Configure logging"""
        log_level = logging.DEBUG if getattr(self.args, 'verbose', False) else logging.INFO
        
        log_dir = os.path.join(self.output_dir, "logs")
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        log_file = os.path.join(log_dir, f'awis_recon_{timestamp}.log')
        
        logging.basicConfig(
            level=log_level,
            format="%(asctime)s %(levelname)s %(message)s",
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )
        self.log = logging.getLogger("awis_core")
        self.log.info(f"Log file: {log_file}")
        
    def ensure_directories(self):
        """Create necessary directories"""
        for directory in [self.output_dir, self.csv_dir, self.text_backup_dir]:
            if not os.path.exists(directory):
                os.makedirs(directory)
                
    def init_ai_clients(self):
        """Initialize AI clients"""
        try:
            api_key = os.getenv('OPENAI_API_KEY')
            if api_key:
                self.openai_client = openai.OpenAI(api_key=api_key)
                self.log.info("ChatGPT API initialized")
        except Exception as e:
            self.log.warning(f"ChatGPT API initialization failed: {e}")
            
        try:
            claude_key = os.getenv('CLAUDE_API_KEY')
            if claude_key:
                self.claude_client = anthropic.Anthropic(api_key=claude_key)
                self.log.info("Claude API initialized")
        except Exception as e:
            self.log.warning(f"Claude API initialization failed: {e}")
    
    def init_csv_file(self, start_url):
        """Initialize CSV file"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        board_name = self.detect_board_name(start_url)
        self.csv_filename = os.path.join(self.csv_dir, f"{board_name}_{timestamp}.csv")
        
        try:
            with open(self.csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(self.csv_headers)
            self.log.info(f"CSV initialized: {self.csv_filename}")
            return True
        except PermissionError:
            self.log.error(f"CSV file locked: {self.csv_filename}")
            return False
            
    def detect_board_name(self, url):
        """Detect board type from URL"""
        url_lower = url.lower()
        if "calcareers" in url_lower:
            return "calcareers"
        elif "ultipro.com" in url_lower or "telecare" in url_lower:
            return "telecare"
        elif "reyes" in url_lower:
            return "reyes"
        elif "edjoin.org" in url_lower:
            return "edjoin"
        elif "usajobs" in url_lower:
            return "usajobs"
        elif "cshs.org" in url_lower or "cedars" in url_lower:
            return "cedarssinai"
        else:
            return "jobs"
            
    def init_webdriver(self, headless=True):
        """Initialize Chrome webdriver"""
        try:
            chrome_options = Options()
            
            if headless:
                chrome_options.add_argument("--headless=new")
                self.log.info("Chrome initialized in headless mode")
            else:
                self.log.info("Chrome initialized in visible mode")
            
            # Core stability flags
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_argument("--ignore-certificate-errors")
            
            # Windows-specific stability flags for Chrome 144+
            chrome_options.add_argument("--disable-software-rasterizer")
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-infobars")
            chrome_options.add_argument("--disable-notifications")
            chrome_options.add_argument("--disable-popup-blocking")
            chrome_options.add_argument("--disable-background-timer-throttling")
            chrome_options.add_argument("--disable-backgrounding-occluded-windows")
            chrome_options.add_argument("--disable-renderer-backgrounding")
            
            chrome_options.page_load_strategy = 'eager'  # Don't wait for full page load

            prefs = {"profile.managed_default_content_settings.images": 2}
            chrome_options.add_experimental_option("prefs", prefs)
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)

            from webdriver_manager.chrome import ChromeDriverManager
            service = Service(ChromeDriverManager().install())

            driver = webdriver.Chrome(service=service, options=chrome_options)
            driver.set_page_load_timeout(60)
            driver.set_script_timeout(60)
            return driver
        except Exception as e:
            self.log.error(f"Failed to initialize Chrome driver: {e}")
            return None
            
    def remove_popups(self):
        """Remove common popups and overlays"""
        for selector in SelectorLibrary.POPUP_SELECTORS:
            try:
                elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                for element in elements:
                    try:
                        if element.is_displayed():
                            self.driver.execute_script("arguments[0].remove();", element)
                    except:
                        continue
            except:
                continue
                
    def try_all_job_link_selectors(self):
        """
        Try every known selector pattern - THE MAGIC
        This is why GMR, United, and unknown sites work
        """
        job_urls = set()
        base_url = f"{urlparse(self.driver.current_url).scheme}://{urlparse(self.driver.current_url).netloc}"
        
        for selector in SelectorLibrary.JOB_LINK_SELECTORS:
            try:
                links = self.driver.find_elements(By.CSS_SELECTOR, selector)
                for link in links:
                    href = link.get_attribute('href')
                    if href:
                        full_url = urljoin(base_url, href)
                        url_lower = full_url.lower()
                        
                        if any(term in url_lower for term in ['job', 'career', 'position', 'opening']):
                            if not any(exc in url_lower for exc in ['search', 'browse', 'login', 'register']):
                                job_urls.add(full_url)
                                
            except Exception:
                continue
        
        job_urls = list(job_urls)
        self.log.info(f"Generic extraction found {len(job_urls)} job URLs")
        return job_urls
    
    def try_all_next_page_selectors(self):
        """
        Try every known pagination pattern
        Returns (element, action) or (None, None)
        """
        # Try CSS selectors
        for selector in SelectorLibrary.NEXT_PAGE_SELECTORS:
            try:
                elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                for element in elements:
                    if element.is_displayed() and element.is_enabled():
                        classes = element.get_attribute('class') or ''
                        if 'disabled' in classes.lower():
                            continue
                        
                        if element.tag_name == 'button':
                            return element, 'BUTTON_CLICK'
                        else:
                            href = element.get_attribute('href')
                            if href and href != self.driver.current_url:
                                return element, href
            except:
                continue
        
        # Try XPath patterns
        for xpath in SelectorLibrary.NEXT_PAGE_XPATHS:
            try:
                elements = self.driver.find_elements(By.XPATH, xpath)
                for element in elements:
                    if element.is_displayed() and element.is_enabled():
                        if element.tag_name == 'button':
                            return element, 'BUTTON_CLICK'
                        else:
                            href = element.get_attribute('href')
                            if href:
                                return element, href
            except:
                continue
        
        return None, None
        
    def extract_text_from_job(self, url):
        """
        Extract text content from job page
        ADAPTIVE: Remembers if JavaScript fallback worked
        """
        try:
            self.driver.get(url)
            time.sleep(1)
            self.remove_popups()
            time.sleep(1)
            
            # If JavaScript mode enabled, skip main selector attempt
            if self.javascript_mode:
                self.log.debug("Using JavaScript extraction (adaptive mode)")
                script = """
                const main = document.querySelector('main');
                if (main && main.innerText && main.innerText.length > 100) {
                    return main.innerText;
                }
                const body = document.body;
                if (body && body.innerText && body.innerText.length > 500) {
                    return body.innerText;
                }
                return "";
                """
                page_text = self.driver.execute_script(script)
                
                if page_text and len(page_text.strip()) > 100:
                    return self.clean_text(page_text)
                else:
                    return None
            
            # Try main selector first
            def main_has_content(driver):
                try:
                    main = driver.find_element(By.TAG_NAME, "main")
                    text = main.text
                    return text and len(text.strip()) > 100
                except:
                    return False
            
            try:
                WebDriverWait(self.driver, 30).until(main_has_content)
                main_element = self.driver.find_element(By.TAG_NAME, "main")
                page_text = main_element.text
                
            except TimeoutException:
                # First timeout - switch to JavaScript mode for all remaining jobs
                self.log.warning("Main element timeout - enabling JavaScript mode for remaining jobs")
                self.javascript_mode = True
                
                script = """
                const main = document.querySelector('main');
                if (main && main.innerText && main.innerText.length > 100) {
                    return main.innerText;
                }
                const body = document.body;
                if (body && body.innerText && body.innerText.length > 500) {
                    return body.innerText;
                }
                return "";
                """
                page_text = self.driver.execute_script(script)
                
                if not page_text or len(page_text.strip()) < 100:
                    return None
            
            if page_text and len(page_text.strip()) > 100:
                return self.clean_text(page_text)
            else:
                raise Exception("Insufficient text")
                
        except Exception as e:
            self.log.error(f"Text extraction failed: {e}")
            return None
    

    def extract_classification_spec(self, job_url):
        """Extract classification specification from CalCareers job posting"""
        try:
            self.driver.get(job_url)
            time.sleep(1)
            
            # Look for classification spec link - STRICT VALIDATION
            # Valid classification spec domains only:
            # - hrnet.calhr.ca.gov/CalHRNet/SpecCrossReference.aspx
            # - eservices.calhr.ca.gov/enterprisehrblazorpublic/Public/ClassSpec/ClassSpecDetail/
            
            classspec_selectors = [
                "a[href*='SpecCrossReference']",      # Most common - hrnet.calhr
                "a[href*='ClassSpecDetail']",          # eservices.calhr
                "a[href*='ClassSpec']",                # Generic classspec
                "a[href*='hrnet.calhr']",              # Domain-specific
                "a[href*='eservices.calhr']"           # Domain-specific
            ]
            
            classspec_xpaths = [
                "//a[contains(@href, 'SpecCrossReference')]",
                "//a[contains(@href, 'ClassSpecDetail')]",
                "//a[contains(@href, 'hrnet.calhr')]",
                "//a[contains(@href, 'eservices.calhr')]",
                "//a[contains(text(), 'Classification')]",
                "//a[contains(text(), 'Class Spec')]"
            ]
            
            # Exclusion patterns - DO NOT accept these URLs
            exclusions = [
                '/Landing/',           # Landing pages
                '/Jobs/',              # Job listing pages
                'NewToStateservice',   # Help pages
                '/Search/',            # Search pages
                'CalHRPublic/Jobs',    # Job posting pages
                'CalHRPublic/Landing'  # Landing pages
            ]
            
            def is_valid_classspec_url(url):
                """Validate URL is actually a classification spec"""
                if not url:
                    return False
                url_lower = url.lower()
                
                # Reject excluded patterns
                for excl in exclusions:
                    if excl.lower() in url_lower:
                        return False
                
                # Must contain one of these valid patterns
                valid_patterns = [
                    'speccrossreference',
                    'classspecdetail',
                    'classspec'
                ]
                
                # Must be from valid domain
                valid_domains = [
                    'hrnet.calhr',
                    'eservices.calhr'
                ]
                
                has_pattern = any(pattern in url_lower for pattern in valid_patterns)
                has_domain = any(domain in url_lower for domain in valid_domains)
                
                return has_pattern and has_domain
            
            classspec_url = None
            attempted_urls = []  # Track what we tried
            
            # Try CSS selectors
            for selector in classspec_selectors:
                try:
                    links = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for link in links:
                        href = link.get_attribute('href')
                        if href:
                            attempted_urls.append(href)
                            if is_valid_classspec_url(href):
                                classspec_url = href
                                self.log.info(f"[OK] Found valid classification spec: {classspec_url}")
                                break
                            else:
                                self.log.debug(f"[X] Rejected URL (invalid): {href}")
                    if classspec_url:
                        break
                except:
                    continue
            
            # Try XPath if CSS failed
            if not classspec_url:
                for xpath in classspec_xpaths:
                    try:
                        links = self.driver.find_elements(By.XPATH, xpath)
                        for link in links:
                            href = link.get_attribute('href')
                            if href:
                                attempted_urls.append(href)
                                if is_valid_classspec_url(href):
                                    classspec_url = href
                                    self.log.info(f"[OK] Found valid classification spec: {classspec_url}")
                                    break
                                else:
                                    self.log.debug(f"[X] Rejected URL (invalid): {href}")
                        if classspec_url:
                            break
                    except:
                        continue
            
            if not classspec_url:
                self.log.warning(f"No classification spec link found. Attempted URLs: {attempted_urls[:3]}")
                return None
            
            # Navigate to classification spec
            self.driver.get(classspec_url)
            time.sleep(3)  # Wait for Blazor JavaScript
            
            # Extract content
            content_selectors = ['main', 'body', '[class*="content"]', '[class*="specification"]']
            
            def validate_classspec_content(text):
                """Verify content is actually a classification spec"""
                if not text or len(text.strip()) < 200:
                    return False
                
                text_lower = text.lower()
                
                # Must contain classification spec indicators
                spec_indicators = [
                    'definition',
                    'minimum qualifications',
                    'knowledge and abilities',
                    'typical tasks',
                    'class code',
                    'schematic'
                ]
                
                # Must NOT contain these (indicates wrong page)
                wrong_page_indicators = [
                    'new to state service',
                    'welcome to',
                    'sign in',
                    'search results',
                    'job postings',
                    'get a state job',
                    'advanced job search',
                    'how to apply for a state job',
                    'find jobs by industry',
                    'exam / assessment search',
                    'geographic job search',
                    'browse jobs listings',
                    'eligible list ranking',
                    'step 1',
                    'step 2',
                    'step 3'
                ]
                
                has_good = sum(1 for ind in spec_indicators if ind in text_lower)
                has_bad = any(ind in text_lower for ind in wrong_page_indicators)
                
                return has_good >= 2 and not has_bad
            
            for selector in content_selectors:
                try:
                    element = self.driver.find_element(By.CSS_SELECTOR, selector)
                    text = element.text
                    if validate_classspec_content(text):
                        self.log.info(f"[OK] Classification spec extracted: {len(text)} chars (validated)")
                        return self.clean_text(text)
                    else:
                        self.log.debug(f"[X] Content rejected: wrong page type")
                except:
                    continue
            
            # JavaScript fallback
            text = self.driver.execute_script("return document.body.innerText;")
            if validate_classspec_content(text):
                self.log.info(f"[OK] Classification spec extracted (fallback): {len(text)} chars")
                return self.clean_text(text)
            
            self.log.warning("Classification spec extraction failed - content validation failed")
            return None
            
        except Exception as e:
            self.log.error(f"Classification spec extraction failed: {e}")
            return None
    
    def extract_calcareers_job_with_classspec(self, url):
        """Extract CalCareers job with classification spec appended"""
        try:
            # Get main job posting
            self.driver.get(url)
            time.sleep(1)
            self.remove_popups()
            time.sleep(1)
            
            # Extract job text using existing logic
            if self.javascript_mode:
                script = """
                const main = document.querySelector('main');
                if (main && main.innerText && main.innerText.length > 100) {
                    return main.innerText;
                }
                return document.body.innerText || "";
                """
                job_text = self.driver.execute_script(script)
            else:
                try:
                    WebDriverWait(self.driver, 30).until(
                        lambda d: d.find_element(By.TAG_NAME, "main").text and len(d.find_element(By.TAG_NAME, "main").text.strip()) > 100
                    )
                    job_text = self.driver.find_element(By.TAG_NAME, "main").text
                except TimeoutException:
                    self.log.warning("Main element timeout - switching to JavaScript mode")
                    self.javascript_mode = True
                    job_text = self.driver.execute_script("return document.querySelector('main').innerText || document.body.innerText;")
            
            if not job_text or len(job_text.strip()) < 100:
                raise Exception("Insufficient job text")
            
            job_text = self.clean_text(job_text)
            
            # Extract classification spec
            classspec_text = self.extract_classification_spec(url)
            
            # Combine
            if classspec_text:
                # Clean classification spec before combining
                classspec_text = self.clean_text(classspec_text)
                classspec_text = self.clean_classification_spec(classspec_text)
                
                combined = f"""{job_text}

========================================
CLASSIFICATION SPECIFICATION (Additional Detail)
========================================

{classspec_text}
"""
                self.log.info(f"Combined text: {len(combined)} chars (Job: {len(job_text)}, ClassSpec: {len(classspec_text)})")
                return combined
            else:
                self.log.info(f"Job text only: {len(job_text)} chars")
                return job_text
                
        except Exception as e:
            self.log.error(f"CalCareers extraction failed: {e}")
            return None

    def clean_text(self, text):
        """Clean unicode characters, mojibake, and remove contact info"""
        import re
        
        cleaned = text.strip()
        
        # Standard unicode replacements
        cleaned = cleaned.replace('\u2019', "'")
        cleaned = cleaned.replace('\u2018', "'")
        cleaned = cleaned.replace('\u201c', '"')
        cleaned = cleaned.replace('\u201d', '"')
        cleaned = cleaned.replace('\u2013', '-')
        cleaned = cleaned.replace('\u2014', '--')
        cleaned = cleaned.replace('\xa0', ' ')
        
        # Mojibake cleaning
        cleaned = cleaned.replace('â€¢', '•')
        cleaned = cleaned.replace('â€"', '-')
        cleaned = cleaned.replace('â€™', "'")
        cleaned = cleaned.replace('â€œ', '"')
        cleaned = cleaned.replace('â€\x9d', '"')
        
        # Remove contact information (case-insensitive, line-by-line)
        lines = cleaned.split('\n')
        cleaned_lines = []
        
        for line in lines:
            line_lower = line.lower()
            
            # Skip lines with contact section headers
            if any(phrase in line_lower for phrase in [
                'human resources contact:',
                'eeo contact:',
                'address for drop-off',
                'address for mailing',
                'department website:'
            ]):
                continue
            
            # Skip lines with phone numbers (various formats)
            if re.search(r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}', line):
                continue
            
            # Skip lines with email addresses (except TBD@workforwarriors.org)
            if re.search(r'[\w\.-]+@[\w\.-]+\.\w+', line):
                if 'tbd@workforwarriors.org' not in line_lower:
                    continue
            
            # Skip physical addresses (street addresses with numbers)
            if re.search(r'\d+\s+[A-Za-z\s]+(?:Street|St|Avenue|Ave|Road|Rd|Boulevard|Blvd|Drive|Dr|Lane|Ln|Court|Ct|Suite|Ste)', line, re.IGNORECASE):
                continue
            
            # Skip ZIP codes standalone
            if re.search(r'\b\d{5}(?:-\d{4})?\b', line) and len(line.strip()) < 30:
                continue
            
            # Skip hours of operation
            if re.search(r'\d{1,2}:\d{2}\s*(?:AM|PM)', line, re.IGNORECASE):
                continue
            
            # Skip California Relay Service / TTY info
            if 'relay service' in line_lower or 'tty' in line_lower:
                continue
            
            # Skip boilerplate phrases
            if any(phrase in line_lower for phrase in [
                'print job save job',
                'equal opportunity employer',
                'you will find additional information about the job in the duty statement'
            ]):
                continue
            
            # Keep the line if it passed all filters
            cleaned_lines.append(line)
        
        return '\n'.join(cleaned_lines)
    
    def clean_classification_spec(self, text):
        """Remove ALL boilerplate from classification specification - AGGRESSIVE"""
        if not text:
            return text
        
        lines = text.split('\n')
        cleaned_lines = []
        
        # Boilerplate phrases to remove (case-insensitive, partial match)
        remove_phrases = [
            'reply_all',
            'return to class spec list',
            'california state personnel board specification',
            'schematic code:',
            'class code:',
            'established:',
            'revised:',
            'title changed:',
            'return to',
            'print',
            'share',
            'back to'
        ]
        
        # Patterns for lines that are ONLY metadata (remove entire line)
        metadata_patterns = [
            r'^\s*\d{1,2}/\d{1,2}/\d{4}\s*$',  # Just dates like "08/07/1980"
            r'^\s*--\s*$',                      # Just dashes
            r'^\s*[A-Z]{2,4}\d{1,3}\s*$',      # Just codes like "JL32" or "4177"
        ]
        
        seen_title = False  # Track if we've seen the main title already
        
        for line in lines:
            line_stripped = line.strip()
            line_lower = line_stripped.lower()
            
            # Skip empty lines at the start (but keep later for formatting)
            if not line_stripped and not cleaned_lines:
                continue
            
            # Skip boilerplate phrases
            if any(phrase in line_lower for phrase in remove_phrases):
                continue
            
            # Skip metadata-only lines
            import re
            if any(re.match(pattern, line_stripped) for pattern in metadata_patterns):
                continue
            
            # Skip duplicate classification title lines
            # Pattern: "Accountant I (Specialist) (4177)"
            # These appear 2-3 times at the top, we want to skip ALL of them
            if '(' in line_stripped and ')' in line_stripped:
                # Check if it looks like a title with classification code
                # Usually has format: "Title Name (Specialty) (Code)"
                paren_count = line_stripped.count('(')
                if paren_count >= 1:
                    # Check if it contains common classification words
                    classification_indicators = [
                        'accountant', 'analyst', 'specialist', 'technician',
                        'officer', 'assistant', 'associate', 'manager',
                        'supervisor', 'coordinator', 'administrator'
                    ]
                    if any(ind in line_lower for ind in classification_indicators):
                        # This looks like a title line - skip it (we don't need it repeated)
                        continue
            
            # If we get here, keep the line
            cleaned_lines.append(line)
        
        # Join and do final cleanup
        cleaned = '\n'.join(cleaned_lines)
        
        # Remove multiple consecutive blank lines (leave max 2)
        while '\n\n\n\n' in cleaned:
            cleaned = cleaned.replace('\n\n\n\n', '\n\n\n')
        
        # Trim leading/trailing whitespace
        cleaned = cleaned.strip()
        
        return cleaned
        
    def apply_wfw_formatting(self, job_data, job_meta):
        """Apply WFW-specific formatting rules"""
        req_id = job_meta.get('req_id', '')
        title = job_data.get('Job Title', '')
        city = job_data.get('City', '')
        location = job_data.get('Location', '')
        
        if title and city:
            job_data['Job Title'] = f"{title} - {city}"
        
        job_data['Job Id'] = req_id
        
        # State extraction from Location field (e.g. "Sacramento, CA")
        if location and ',' in location:
            parts = location.split(',')
            if len(parts) >= 2:
                job_data['State'] = parts[-1].strip()
        elif city and ',' in city:
            # Fallback: extract from City field
            parts = city.split(',')
            if len(parts) >= 2:
                job_data['State'] = parts[-1].strip()
        
        # Country default
        if not job_data.get('Country'):
            job_data['Country'] = 'United States'
        
        # Zip code extraction from Address or Location
        import re
        address = job_data.get('Address', '')
        if address:
            zip_match = re.search(r'\b\d{5}(?:-\d{4})?\b', address)
            if zip_match:
                job_data['Zip Code'] = zip_match.group(0)
        if not job_data.get('Zip Code') and location:
            zip_match = re.search(r'\b\d{5}(?:-\d{4})?\b', location)
            if zip_match:
                job_data['Zip Code'] = zip_match.group(0)
        
        # Defaults
        job_data['Apply Email'] = 'TBD@workforwarriors.org'
        job_data['Full Name'] = 'TBD'
        job_data['Apply URL'] = ''
        
        if job_data.get('Employer Email'):
            job_data['Employer Email'] = f"999{job_data['Employer Email']}"
        
        return job_data
    
    def robust_json_repair(self, raw_response):
        """
        Three-layer JSON repair: Extract from noise, fix truncation, targeted mojibake cleanup
        """
        try:
            # Layer 1: Extract JSON from conversational noise
            match = re.search(r'(\{.*\})', raw_response, re.DOTALL)
            if not match:
                raise ValueError("No JSON braces found in response")
            
            clean_str = match.group(1)
            
            # Layer 2: Fix truncation (missing closing braces)
            open_braces = clean_str.count('{')
            close_braces = clean_str.count('}')
            if open_braces > close_braces:
                clean_str += '}' * (open_braces - close_braces)
            
            # Layer 3: Targeted mojibake replacement (not aggressive removal)
            mojibake_map = {
                'â€¢': '•',
                'â€"': '—',
                'â€™': "'",
                'â€œ': '"',
                'â€': '"',
                'Â': '',
            }
            for bad, good in mojibake_map.items():
                clean_str = clean_str.replace(bad, good)
            
            # Attempt parse
            data = json.loads(clean_str)
            return data
            
        except Exception as e:
            self.log.error(f"JSON repair failed: {e}")
            raise
    
    def parse_with_chatgpt(self, text_content, job_meta):
        """Parse job text using ChatGPT API"""
        if not self.openai_client:
            raise ProviderUnavailableError("ChatGPT client not available")
        
        # LAYER 1: PRE-PROCESS INPUT - Escape quotes to prevent JSON breaking
        # Replace double quotes with escaped version so ChatGPT never sees unescaped quotes
        text_content_safe = text_content.replace('"', '\\"')
            
        try:
            prompt = f"""STRICT OUTPUT REQUIREMENT: Return ONLY raw JSON. No conversational text. No markdown backticks. No preamble. No postscript. Response must start with {{ and end with }}.

OBJECTIVE: Extract CalCareers job data into exactly 22 fields.

CRITICAL JSON FORMATTING RULES:
1. Output MUST be valid, parseable JSON
2. ALL string values MUST escape double quotes with backslash
3. Example WRONG: "title": "Clinical Social Worker- (ISUDT), "level of care", Medical"
4. Example RIGHT: "title": "Clinical Social Worker- (ISUDT), \\"level of care\\", Medical"
5. Check EVERY field for embedded quotes before returning
6. Apostrophes/single quotes are OK and don't need escaping

REAL FAILURE EXAMPLES TO AVOID:
❌ BREAKS: "description": "Provides "high quality" services"
✓ WORKS: "description": "Provides \\"high quality\\" services"

❌ BREAKS: "qualifications": "Must have "strong communication" skills"
✓ WORKS: "qualifications": "Must have \\"strong communication\\" skills"

CLASSIFICATION SPEC FILTERING:
- If classification spec is present, extract ONLY these sections: "Minimum Qualifications", "Knowledge and Abilities", "Definition"
- REMOVE all boilerplate: "reply_all", "Return to Class Spec List", "California State Personnel Board Specification"
- REMOVE metadata lines: "Schematic Code:", "Class Code:", "Established:", "Revised:", "Title Changed:"
- If classification spec URL is a landing/help page (not actual job requirements), IGNORE it and process only main job posting
- NO-FAIL CONSTRAINT: Never let bad classification spec cause JSON truncation

JOB POSTING TEXT:
{text_content_safe}

REQUIRED JSON SCHEMA (22 FIELDS):

{{
  "Job Title": "JC-##### - [Classification] - [Working Title]",
  "Job Description": "[Complete job posting + filtered classification spec]",
  "Job Type": "[Full-time/Part-time/etc.]",
  "Categories": "[Category 1]; [Category 2]; [Category 3]",
  "Location": "[City, State]",
  "City": "[City name only]",
  "State": "[State abbreviation - CA, NY, etc.]",
  "Country": "United States",
  "Zip Code": "[5-digit zip if present]",
  "Address": "[Full street address if present]",
  "Salary From": "[Number only - no $ or commas]",
  "Salary To": "[Number only - no $ or commas]",
  "Salary Period": "[Hourly/Monthly/Yearly]",
  "Apply Email": "TBD@workforwarriors.org",
  "Posting Date": "[YYYY-MM-DD format if present]",
  "Expiration Date": "[YYYY-MM-DD format if present]",
  "Qualifications": "[Max 250 chars - education, experience, licenses]",
  "Employer Email": "[Department email if present or empty]",
  "Full Name": "TBD",
  "Company Name": "State of California - [Department Name]",
  "Employer Website": "[Department website if present or empty]",
  "Employer Phone": "[Department phone if present or empty]"
}}

FIELD EXTRACTION DETAILS:

Job Title:
- Format: "JC-[number] - [Classification] - [Working Title]"
- Example: "JC-505028 - ACCOUNTANT I (SPECIALIST) - Check Processor"
- Append Job Id/Control number to title

Job Description:
- Include EVERY WORD from job posting
- If classification spec present, append filtered version (Minimum Qualifications, Knowledge and Abilities, Definition sections only)
- Remove all boilerplate and metadata from classification spec

Qualifications:
- Extract complete minimum qualifications
- If exceeds 250 chars, intelligently truncate to 200-245 chars
- Priority: Education → Experience → Licenses → Top 2-3 skills
- Use semicolons to separate items

Categories:
- Select EXACTLY 3 most relevant from this list: Accounting & Auditing; Administrative & Clerical; Government; Health Care; Human Resources; Information Technology; Legal; Management; Nonprofit - Social Services; Professional Services; Public Relations; Science; Government - Federal; Government - Local; Government - State

Salary From/To:
- Numbers only, no $ or commas
- "$50,000" becomes "50000"
- "50k" becomes "50000"

Country:
- Always "United States"

Apply Email:
- Always "TBD@workforwarriors.org"

Full Name:
- Always "TBD"

MANDATORY: Return ONLY the JSON object. No text before or after.

ALL OTHER FIELDS: Extract if present, leave empty ("") if not found

RULES:
- Use exact field names from template
- Replace commas in values with semicolons
- Leave fields empty ("") if not found in posting
- NO PLACEHOLDERS - real data only or blank
- Job Description gets EVERYTHING - no summarization

Return valid JSON with all {len(self.csv_headers)} fields.
"""

            response = self.openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
                temperature=0,
                max_tokens=4000
            )
            
            result = response.choices[0].message.content.strip()
            
            # LAYER 2: POST-PROCESS OUTPUT - Use robust JSON repair
            # Handles conversational preambles, markdown fences, truncation, mojibake
            try:
                job_data = self.robust_json_repair(result)
            except Exception as repair_error:
                # Robust repair failed - log and trigger backup mechanism
                self.log.warning(f"JSON repair failed: {repair_error}")
                raise json.JSONDecodeError(f"JSON repair failed: {repair_error}", result, 0)
                
            # LAYER 3: APPLY DEFAULTS - Ensure required fields are populated
            job_data = self.apply_wfw_formatting(job_data, job_meta)
            
            
            return job_data
            
        except json.JSONDecodeError as e:
            self.log.error(f"ChatGPT returned invalid JSON: {e}")
            self.api_failures += 1
            raise ProviderUnavailableError("Invalid JSON response")
        except Exception as e:
            error_str = str(e).lower()
            if "quota" in error_str or "429" in error_str or "rate_limit" in error_str:
                self.api_failures += 1
                raise QuotaExceededError(f"ChatGPT quota exceeded: {e}")
            self.api_failures += 1
            raise ProviderUnavailableError(f"ChatGPT error: {e}")
    
    def parse_with_claude(self, text_content, job_meta):
        """Parse job text using Claude API"""
        if not self.claude_client:
            raise ProviderUnavailableError("Claude client not available")
            
        try:
            prompt = f"""Extract structured job data from this CalCareers posting. Return ONLY valid JSON with exactly these {len(self.csv_headers)} fields.

CRITICAL INSTRUCTIONS:
1. CAPTURE EVERY WORD from the job posting - do not summarize or condense
2. Job Description field is UNLIMITED - include the complete posting verbatim
3. Qualifications field has 250 char limit - extract FULL requirements first, then intelligently truncate
4. CLEAN TEXT: Remove mojibake characters (â€¢ → bullet, â€" → dash)

JOB POSTING TEXT:
{text_content}

FIELD EXTRACTION RULES:

Job Title: 
Combine in format: "JC-##### - [Classification] - [Working Title]"
- Extract JC number from "JC-######" or "Job Control"  
- Extract Classification (e.g. "ACCOUNTANT I (SPECIALIST)")
- Extract Working Title (e.g. "Check Processor")
- Example: "JC-505028 - ACCOUNTANT I (SPECIALIST) - Check Processor"

Job Description: 
- Include EVERY WORD from the job posting exactly as written
- Include all sections: duties, responsibilities, requirements, benefits, EOE statements
- Include all application instructions and deadlines
- Include all contact information
- Include all program details and shift information
- DO NOT EXCLUDE ANYTHING from the original posting
- If classification specification text is present (marked with separator line), include it COMPLETELY at the end
- Format: [Complete original job posting]\\n\\n[Complete classification spec if present]

Qualifications:
STEP 1 - Extract COMPLETE minimum qualifications including:
- Full education requirement (every word)
- Full experience requirement (every word)  
- All licenses/certifications (every word)
- All required skills (every word)

STEP 2 - If total exceeds 250 characters, intelligently truncate:
- Keep education requirement (abbreviated if needed: "Bachelor's in [field]")
- Keep experience requirement (abbreviated: "X years [type]")
- Keep critical licenses (abbreviated: "[License] required")
- Keep top 2-3 skills only
- Use semicolons to separate
- Target: 200-245 characters (leave buffer for safety)

EXAMPLE QUALIFICATIONS TRUNCATION:
FULL (320 chars): "Education: Equivalent to graduation from a four-year college with major study in accounting, business administration with an emphasis in accounting, or a closely related field. Experience: Two years of professional accounting experience. License: Possession of a valid California CPA certificate."

TRUNCATED (238 chars): "Bachelor's degree in Accounting or Business Administration with accounting emphasis; 2 years professional accounting experience; Valid California CPA certificate required; Knowledge of GAAP and governmental accounting"

Location: Extract as "City, State" format (e.g. "Sacramento, CA")

City: Extract city name only (e.g. "Sacramento")

State: Extract state abbreviation (e.g. "CA", "NY", "TX")

Country: Leave blank (will be auto-populated to "United States")

Zip Code: Extract 5-digit zip code if present in address or location text

Address: Extract full street address if present (e.g. "1501 Capitol Avenue, Suite 71.1501, Sacramento, CA 95814")

Salary From/To: Extract as numbers only, remove $ and commas
- "$50,000" becomes "50000"
- "50k" becomes "50000"

Salary Period: "Hourly", "Monthly", or "Yearly"

Job Type: Extract exact type (Full-time, Part-time, etc.) with hours if mentioned

Categories: Select EXACTLY 3 most relevant categories separated by semicolons from: Accounting & Auditing; Administrative & Clerical; Advertising & Marketing; Agriculture; Architecture & Design; Art & Creative; Automotive; Banking; Biotech; Business Development; Construction; Consultant; Customer Service; Defense & Military; Distribution - Shipping; Education; Engineering; Entry Level; Executive; Facilities; Finance; Franchise; General Business; General Labor; Government; Grocery; Health Care; Hospitality - Hotel; Human Resources; Information Technology; Installation-Maint-Repair; Insurance; Inventory; Janitorial; Legal; Legal Admin; Logistics; Management; Manufacturing; Media - Journalism; Nonprofit - Social Services; Nurse; Other; Pharmaceutical; Professional Services; Project Management; Public Relations; Purchasing - Procurement; QA - Quality Control; Real Estate; Research; Restaurant - Food Service; Retail; Sales; Science; Skilled Labor; Strategy - Planning; Supply Chain; Telecommunications; Training; Transportation; Travel; Warehouse; Web Design; Web Development; Writing - Editing; Energy; Environmental; Entertainment; Fashion; Fitness; Fundraising; Gaming; Government - Federal; Government - Local; Government - State; Home Services; Landscaping; Parenting; Pet Care; Security; Sports

Company Name: "State of California - [Department Name]" or specific hiring department

Company Description: 2-3 sentences about the department/agency from "About" sections

Posting Date: Extract in YYYY-MM-DD format if present

Expiration Date: Extract application deadline in YYYY-MM-DD format if present

Apply Email: Leave blank (will be auto-populated)

Apply URL: Leave blank (will be auto-populated)

ALL OTHER FIELDS: Extract if present, leave empty ("") if not found

RULES:
- Use exact field names from template
- Replace commas in values with semicolons
- Leave fields empty ("") if not found in posting
- NO PLACEHOLDERS - real data only or blank
- Job Description gets EVERYTHING - no summarization

Return valid JSON with all {len(self.csv_headers)} fields.
"""

            response = self.claude_client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=4000,
                temperature=0,
                messages=[{"role": "user", "content": prompt}]
            )
            
            result = response.content[0].text.strip()
            
            # Strip code fences
            json_match = re.search(r'```json\s*(\{.*?\})\s*```', result, re.DOTALL)
            if json_match:
                result = json_match.group(1)
            elif result.startswith('```json'):
                result = result[7:].rstrip('```').strip()
            elif result.startswith('```'):
                result = result[3:].rstrip('```').strip()
                
            job_data = json.loads(result.strip())
            job_data = self.apply_wfw_formatting(job_data, job_meta)
            
            return job_data
            
        except json.JSONDecodeError as e:
            self.log.error(f"Claude returned invalid JSON: {e}")
            self.api_failures += 1
            raise ProviderUnavailableError("Invalid JSON response")
        except Exception as e:
            error_str = str(e).lower()
            if "quota" in error_str or "429" in error_str:
                self.api_failures += 1
                raise QuotaExceededError(f"Claude quota exceeded: {e}")
            self.api_failures += 1
            raise ProviderUnavailableError(f"Claude error: {e}")
    
    def save_backup_artifact(self, content, url):
        """Save text backup artifact"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            req_id = SelectorLibrary.extract_req_id(url)
            
            filename = os.path.join(self.text_backup_dir, f"job_{req_id}_{timestamp}.txt")
            backup_content = f"""JOB TEXT BACKUP
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
URL: {url}
REQ ID: {req_id}

--- CONTENT ---
{content}
"""
            
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(backup_content)
                
            self.text_backups_saved += 1
            return filename
            
        except Exception as e:
            self.log.error(f"Backup save failed: {e}")
            return None
    
    def save_job_to_csv(self, job_data_dict):
        """Append job data to CSV file"""
        try:
            if self.job_counter % 10 == 0:
                try:
                    with open(self.csv_filename, 'a', encoding='utf-8') as test:
                        pass
                except PermissionError:
                    self.log.error("CSV became locked mid-run")
                    raise
            
            with open(self.csv_filename, 'a', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile, quoting=csv.QUOTE_MINIMAL)
                csv_row = [job_data_dict.get(header, "") for header in self.csv_headers]
                writer.writerow(csv_row)
            return True
        except Exception as e:
            self.log.error(f"CSV save failed: {e}")
            return False
    
    def process_job(self, job_url, preview_callback=None):
        """Process single job with optional preview callback"""
        req_id = SelectorLibrary.extract_req_id(job_url)
        
        if req_id and req_id in self.processed_req_ids:
            return True
        
        self.job_counter += 1
        self.log.info(f"Processing [{self.job_counter}] REQ {req_id}")
        
        # CalCareers: Extract with classification spec
        if "calcareers" in job_url.lower():
            content = self.extract_calcareers_job_with_classspec(job_url)
        else:
            content = self.extract_text_from_job(job_url)
            
        if not content or len(str(content).strip()) < 50:
            self.log.warning("No content extracted")
            self.failed_captures += 1
            
            # Call preview callback even on failure
            if preview_callback:
                preview_callback(None, req_id, False)
            
            return False
            
        saved_backup = None
        def ensure_backup():
            nonlocal saved_backup
            if saved_backup is None:
                saved_backup = self.save_backup_artifact(content, job_url)
                
        if getattr(self.args, 'pipeline', 'api') == "backup":
            ensure_backup()
            if req_id:
                self.processed_req_ids.add(req_id)
                self.save_processed_id(req_id)
            self.successful_captures += 1
            
            if preview_callback:
                preview_callback(None, req_id, True)
            
            return True
            
        if getattr(self.args, 'backup_policy', 'onfail') == "always":
            ensure_backup()
            
        ai_provider = getattr(self.args, 'ai', 'chatgpt')
        
        parsed_ok = False
        job_data = None
        try:
            job_meta = {"req_id": req_id, "url": job_url}
            
            if ai_provider == 'chatgpt':
                job_data = self.parse_with_chatgpt(content, job_meta)
            elif ai_provider == 'claude':
                job_data = self.parse_with_claude(content, job_meta)
            else:
                job_data = None
            
            if job_data:
                if self.save_job_to_csv(job_data):
                    parsed_ok = True
                    self.successful_captures += 1
                    self.log.info(f"Job {self.job_counter} parsed via {ai_provider}")
                    
        except QuotaExceededError as e:
            self.log.warning(f"Provider quota exceeded: {e}")
            self.api_failures += 1
        except ProviderUnavailableError as e:
            self.log.warning(f"Provider unavailable: {e}")
            self.api_failures += 1
            
        if not parsed_ok and getattr(self.args, 'backup_policy', 'onfail') == "onfail":
            ensure_backup()
            
        if req_id:
            self.processed_req_ids.add(req_id)
            self.save_processed_id(req_id)
        
        if not parsed_ok:
            self.failed_captures += 1
        
        # Call preview callback with job data
        if preview_callback:
            preview_callback(job_data, req_id, parsed_ok)
            
        return parsed_ok or (saved_backup is not None)
    
    def load_processed_ids(self):
        """Load previously processed job IDs"""
        id_file = os.path.join(self.output_dir, 'processed_job_ids.txt')
        if os.path.exists(id_file):
            with open(id_file, 'r') as f:
                return set(line.strip() for line in f if line.strip())
        return set()

    def save_processed_id(self, req_id):
        """Append processed job ID to persistent file"""
        id_file = os.path.join(self.output_dir, 'processed_job_ids.txt')
        with open(id_file, 'a') as f:
            f.write(f"{req_id}\n")
    
    def run_with_auto_pagination(self, start_url, max_pages=50):
        """
        Universal scraper with smart auto-pagination
        Works on ANY site - GMR, United, CalCareers, unknown employers
        """
        self.log.info("AWIS Core - Starting automation")
        self.log.info(f"Pipeline: {getattr(self.args, 'pipeline', 'api')}")
        self.log.info(f"AI: {getattr(self.args, 'ai', 'chatgpt')}")
        self.log.info(f"Backup Policy: {getattr(self.args, 'backup_policy', 'onfail')}")
        
        if not self.init_csv_file(start_url):
            return False
        
        self.driver = self.init_webdriver(headless=True)
        if not self.driver:
            self.log.error("Failed to initialize driver")
            return False
            
        try:
            self.driver.get(start_url)
            time.sleep(2)
            
            current_page = 1
            consecutive_failures = 0
            
            while current_page <= max_pages:
                self.log.info(f"=== PAGE {current_page} ===")
                
                # Extract jobs using generic selectors
                job_urls_before = len(self.processed_req_ids)
                job_urls = self.try_all_job_link_selectors()
                
                if not job_urls:
                    self.log.info("No jobs found on page")
                    consecutive_failures += 1
                    if consecutive_failures >= 2:
                        self.log.info("Two empty pages - stopping")
                        break
                else:
                    consecutive_failures = 0
                
                # Process jobs
                for job_url in job_urls:
                    self.process_job(job_url)
                    time.sleep(2)
                
                # Try pagination
                next_element, next_action = self.try_all_next_page_selectors()
                
                if not next_element:
                    self.log.info("No more pages detected")
                    break
                
                try:
                    if next_action == 'BUTTON_CLICK':
                        self.driver.execute_script("arguments[0].click();", next_element)
                        self.log.info("Clicked next page button")
                    else:
                        self.driver.get(next_action)
                        self.log.info(f"Navigated to: {next_action}")
                    
                    time.sleep(2)
                    
                    # Verify we're on a new page
                    job_urls_after = len(self.processed_req_ids)
                    if job_urls_after == job_urls_before:
                        self.log.warning("No new jobs after pagination")
                        consecutive_failures += 1
                        if consecutive_failures >= 2:
                            break
                    
                except Exception as e:
                    self.log.error(f"Navigation error: {e}")
                    break
                
                current_page += 1
                
            self.log.info(f"Pagination complete: {current_page} pages processed")
            
        except KeyboardInterrupt:
            self.log.info("Automation stopped by user")
        except Exception as e:
            self.log.error(f"Automation error: {e}")
        finally:
            self.cleanup()
            
        return True
    
    def cleanup(self):
        """Cleanup and summary"""
        if self.driver:
            self.driver.quit()
            
        duration = datetime.now() - self.session_start_time
        
        self.log.info("=== RECON COMPLETE ===")
        self.log.info(f"Jobs processed: {self.job_counter}")
        self.log.info(f"Successful: {self.successful_captures}")
        self.log.info(f"Failed: {self.failed_captures}")
        self.log.info(f"API failures: {self.api_failures}")
        self.log.info(f"Backups saved: {self.text_backups_saved}")
        self.log.info(f"Runtime: {duration}")
        if self.job_counter > 0:
            success_rate = (self.successful_captures / self.job_counter) * 100
            self.log.info(f"Success rate: {success_rate:.1f}%")
