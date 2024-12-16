import pandas as pd
import requests
import xml.etree.ElementTree as ET
from tkinter import Tk, filedialog, Button, Label, messagebox
import os
from urllib.parse import urlparse
from io import StringIO
import csv
from requests.auth import HTTPBasicAuth
from datetime import datetime

#To use bigger limit on the buffer and stack to avoid overflow
csv.field_size_limit(10**6)

class TokenAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Token Analyzer")
        self.root.geometry("400x200")
        
        self.csv_path = None
        self.output_path = None

        # Labels and Buttons
        self.label = Label(root, text="Load Ahrefs CSV File")
        self.label.pack(pady=10)

        self.upload_button = Button(root, text="Browse", command=self.load_csv)
        self.upload_button.pack(pady=5)

        self.generate_button = Button(root, text="Generate Report", command=self.generate_report)
        self.generate_button.pack(pady=5)

        self.clear_button = Button(root, text="Clear All", command=self.clear_all)
        self.clear_button.pack(pady=5)

        self.close_button = Button(root, text="Exit Program", command=root.quit)
        self.close_button.pack(pady=5)

    def load_csv(self):
        self.csv_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.csv_path:
            messagebox.showinfo("File Loaded", f"File loaded: {self.csv_path}")
        else:
            messagebox.showerror("Error", "No file selected.")

    def extract_tokens(self, url):
        try:
            start = url.find("_") + 1
            end = url.find("/", start)
            token = url[start:end]
            if len(token) == 32:
                return token
        except Exception as e:
            print(f"Error extracting token from {url}: {e}")
        return None

    def fetch_api_data(self, tokens):
        url = f"https://admin.throneneataffiliates.com/feeds.php?FEED_ID=4&TOKENS={','.join(tokens)}"
        try:
            print(f"Sending tokens to API: {tokens}")
            response = requests.get(url, auth=HTTPBasicAuth('TokenAPI', 'ToKenChAnGeMePaS$1234'))
            if response.status_code == 200:
                return response.text
            elif response.status_code == 401:
                print("Error 401: Authorization failed. Please check credentials.")
            else:
                print(f"Error connecting to API for tokens: {response.status_code}")
        except Exception as e:
            print(f"Error sending request for tokens: {e}")
        return None

    def parse_xml(self, xml_data):
        data = {}
        try:
            root = ET.fromstring(xml_data)
            for token_element in root.findall(".//TOKEN"):
                token = token_element.get("PREFIX")
                setup_element = token_element.find(".//SETUP")
                user_element = token_element.find(".//USER")
                data[token] = {
                    'username': user_element.get("USERNAME") if user_element is not None else '',
                    'access_url': f"https://admin.throneneataffiliates.com/affiliate_summary.php?id={setup_element.get('OBJECT_ID')}" if setup_element is not None else '',
                    'object_description': setup_element.get("OBJECT_DESCRIPTION") if setup_element is not None else ''
                }
        except Exception as e:
            print(f"Error parsing XML: {e}")
        return data

    def generate_report(self):
        if not self.csv_path:
            messagebox.showerror("Error", "Please load a CSV file first.")
            return

        try:
            # Read the file as raw text and handle encoding errors
            with open(self.csv_path, 'r', encoding='utf-16', errors='ignore') as file:
                first_line = file.readline()
                print(f"First line of the file: {first_line}")
                delimiter = ',' if ',' in first_line else ';' if ';' in first_line else '\t'

            with open(self.csv_path, 'r', encoding='utf-16', errors='ignore') as file:
                raw_data = file.read()

            # Load the raw data into pandas using the detected delimiter
            df = pd.read_csv(StringIO(raw_data), delimiter=delimiter, on_bad_lines='skip', engine='python')

            # Clean column names
            df.columns = df.columns.str.strip().str.lower()

            # Check for necessary columns
            if 'target url' not in df.columns or 'referring page url' not in df.columns:
                raise ValueError("The CSV file does not contain the required columns 'Target URL' or 'Referring Page URL'.")

            # Extract all unique tokens
            df['token'] = df['target url'].apply(self.extract_tokens)
            all_tokens = df['token'].dropna().unique().tolist()
            print(f"Unique tokens extracted: {all_tokens}")

            # Send all tokens to the API in a single request
            xml_data = self.fetch_api_data(all_tokens)

            # Parse the API response
            api_data = self.parse_xml(xml_data)

            # Build data row by row
            report_data = []
            for _, row in df.iterrows():
                token = row['token']
                referring_url = row['referring page url']
                target_url = row['target url']
                last_seen = row.get('last seen', 'N/A')

                token_data = api_data.get(token, {
                    'username': '',
                    'access_url': '',
                    'object_description': ''
                })

                # Add row to report
                report_data.append({
                    'Affiliate': f'=HYPERLINK("{token_data["access_url"]}", "{token_data["username"]}")' if token_data["username"] else '',
                    'Website': referring_url,
                    'Landing Page': token_data['object_description'],
                    'Tracking Link': target_url,
                    'Date': last_seen
                })

            # Generate DataFrame and save as Excel
            report_df = pd.DataFrame(report_data)
            print(f"Final DataFrame:\n{report_df}")

            # Generate file name with current date
            first_url = df['target url'].iloc[0]
            domain = urlparse(first_url).netloc.replace(".", "_")
            current_date = datetime.now().strftime("%d-%m-%Y")
            self.output_path = os.path.expanduser(f"~/Desktop/{domain}_Ahrefs_Report_{current_date}.xlsx")
            report_df.to_excel(self.output_path, index=False)

            messagebox.showinfo("Success", f"Report generated: {self.output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error generating the report: {e}")

    def clear_all(self):
        self.csv_path = None
        if self.output_path and os.path.exists(self.output_path):
            os.remove(self.output_path)
        messagebox.showinfo("Reset", "Data cleared. Ready to start again.")

# Create the application
root = Tk()
app = TokenAnalyzerApp(root)
root.mainloop()
