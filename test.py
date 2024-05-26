import os
import pandas as pd
import pyodbc as odbc
import time
#import pandas as pd
from underthesea import sentiment
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException



# read by default 1st sheet of an excel file
dataframe1 = pd.read_excel('Data.xlsx')
#print(dataframe1)

def take_URL(df):
    urls =[]
    for URL in df['Link']:
        urls.append(URL)
    return urls

take_URL(dataframe1)



def remove_first_two_sentences(paragraph):
    # Split the paragraph into sentences based on newline character
    sentences = paragraph.split('\n')
    
    # Remove the first two sentences
    remaining_sentences = sentences[2:]
    
    # Join the remaining sentences back into a paragraph
    remaining_paragraph = '\n'.join(remaining_sentences)
    
    return remaining_paragraph

def get_sentiment(text):
    return sentiment(text)

def crawl_comments(url):
    chrome_path = "D:/Audition/chromedriver.exe"
    chrome_options = webdriver.ChromeOptions()
    chrome_options.binary_location = "C:/Program Files/Google/Chrome/Application/chrome.exe"
    driver = webdriver.Chrome(options=chrome_options)
    driver.get(url)
    time.sleep(5)

    # Scroll down to load more comments (you might need to adjust this based on the website)
    for i in range(5):  # Increase the range to scroll more
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.END)
        time.sleep(2)

    comments_list = []

    # Initial loop to load comments from the first page
    while True:
        comments = driver.find_elements(By.XPATH, '//*[@id="module_product_review"]//div[@class="item"]')
        for comment in comments:
            commentText = remove_first_two_sentences(comment.text)
            #print(commentText)
            comments_list.append(commentText)

        # Try to find the next page button
        try:
            next_button = driver.find_element(By.XPATH, '//button[@class="next-btn"]')
            if next_button.is_enabled():  # Check if the next button is clickable
                next_button.click()
                time.sleep(2)
            else:
                break  # Exit the loop if the next button is disabled (no more pages)
        except NoSuchElementException:
            break  # Exit the loop if the next button is not found (no more pages)

    driver.quit()
    return comments_list

url_to_crawl = take_URL(dataframe1)
#all_comments = []
comments_dict_with_sentiment = {}
comments_dict_without_sentiment = {}

for url in url_to_crawl:
    comments = crawl_comments(url)
    #all_comments.extend(comments)
    df_comments = pd.DataFrame(comments, columns=['Comments'])
    df_comments['Sentiment'] = df_comments['Comments'].apply(get_sentiment)
    if url in comments_dict_with_sentiment:
        # Append the new comments to the existing DataFrame
        comments_dict_with_sentiment[url] = pd.concat([comments_dict_with_sentiment[url], df_comments], ignore_index=True)
    else:
        # Add the new DataFrame to the dictionary
        comments_dict_with_sentiment[url] = df_comments
    df_comments_without_sentiment = df_comments.drop(columns=['Sentiment'])
    if url in comments_dict_without_sentiment:
        comments_dict_without_sentiment[url] = pd.concat([comments_dict_without_sentiment[url], df_comments_without_sentiment], ignore_index=True)
    else:
        comments_dict_without_sentiment[url] = df_comments_without_sentiment



# for url, df_comments in comments_dict_without_sentiment.items():
#     print(f"URL: {url}")
#     print(df_comments)
#     print("-" * 50)


DRIVER_NAME = 'SQL SERVER'
SERVER_NAME = 'DESKTOP-E7APER1\SQLEXPRESS'
DATABASE_NAME = 'JJ'
conection = f"""
    DRIVER={{{DRIVER_NAME}}};
    SERVER={SERVER_NAME};
    DATABASE={DATABASE_NAME};
    Trust_Connection = yes;
"""

conn = odbc.connect(conection)
cursor = conn.cursor()
print(conn)

cursor = conn.cursor()
cursor.execute("DROP TABLE IF EXISTS chai_nuoc")
conn.commit()

cursor = conn.cursor()
cursor.execute("DROP TABLE IF EXISTS ao_thun")
conn.commit()

cursor = conn.cursor()
cursor.execute("DROP TABLE IF EXISTS kep_toc")
conn.commit()

cursor = conn.cursor()
cursor.execute("DROP TABLE IF EXISTS loa_bluetooth")
conn.commit()

cursor = conn.cursor()
cursor.execute("DROP TABLE IF EXISTS op_dien_thoai")
conn.commit()
#################################################################################
#Create the table
cursor.execute("""
CREATE TABLE chai_nuoc (
    comment NVARCHAR(1000),
    Sentiment NVARCHAR(200)
)
""")
conn.commit()

cursor.execute("""
CREATE TABLE ao_thun (
    comment NVARCHAR(1000),
    Sentiment NVARCHAR(200)
)
""")
conn.commit()

cursor.execute("""
CREATE TABLE kep_toc (
    comment NVARCHAR(1000),
    Sentiment NVARCHAR(200)
)
""")
conn.commit()

cursor.execute("""
CREATE TABLE loa_bluetooth (
    comment NVARCHAR(1000),
    Sentiment NVARCHAR(200)
)
""")
conn.commit()

cursor.execute("""
CREATE TABLE op_dien_thoai (
    comment NVARCHAR(1000),
    Sentiment NVARCHAR(200)
)
""")
conn.commit()



def get_desired_key_for_dataframe(url):
    if "chai-nuoc" in url:
        return "chai_nuoc"
    elif "kep-toc" in url:
        return "kep_toc"
    elif "op-dien-thoai" in url:
        return "op_dien_thoai"
    elif "ao-thun" in url:
        return "ao_thun"
    else:
        return "loa_bluetooth"

# Create a new dictionary with updated keys
new_comments_dict = {}

# Iterate through the items in the original dictionary
for url, df_comments in comments_dict_with_sentiment.items():
    # Generate the desired key for the DataFrame
    desired_key = get_desired_key_for_dataframe(url)
    
    # Assign the DataFrame to the new key in the new dictionary
    new_comments_dict[desired_key] = df_comments

    # Optionally, print the information
    print(f"{desired_key}")
    print(df_comments)
    print("-" * 50)


df_aothun = new_comments_dict["ao_thun"]
#print(df_aothun)

for index, row in df_aothun.iterrows():
    cursor.execute("INSERT INTO ao_thun (comment, Sentiment) VALUES (?, ?)", row['Comments'], row['Sentiment'])

conn.commit()

df_chainuoc = new_comments_dict["chai_nuoc"]
#print(df_chainuoc)

for index, row in df_chainuoc.iterrows():
    cursor.execute("INSERT INTO chai_nuoc (comment, Sentiment) VALUES (?, ?)", row['Comments'], row['Sentiment'])

conn.commit()

df_keptoc = new_comments_dict["kep_toc"]
#print(df_keptoc)

for index, row in df_keptoc.iterrows():
    cursor.execute("INSERT INTO kep_toc (comment, Sentiment) VALUES (?, ?)", row['Comments'], row['Sentiment'])

conn.commit()

df_opdt = new_comments_dict["op_dien_thoai"]
#print(df_opdt)

for index, row in df_opdt.iterrows():
    cursor.execute("INSERT INTO op_dien_thoai (comment, Sentiment) VALUES (?, ?)", row['Comments'], row['Sentiment'])

conn.commit()

df_loa = new_comments_dict["loa_bluetooth"]
#print(df_loa)

for index, row in df_loa.iterrows():
    cursor.execute("INSERT INTO loa_bluetooth (comment, Sentiment) VALUES (?, ?)", row['Comments'], row['Sentiment'])

conn.commit()


#################################
json_dataframes = {key: df.to_json(orient='split') for key, df in new_comments_dict.items()}

# HTML template
html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>DataFrame Viewer</title>
    <style>
        .button-55 {
            padding: 10px 20px;
            font-size: 16px;
            cursor: pointer;
            background-color: #4CAF50;
            color: white;
            border: none;
            border-radius: 5px;
            margin: 5px;
        }
    </style>
</head>
<body>
    <div id="buttons">
        [button]
    </div>
    <div id="display">
        <!-- DataFrame content will be displayed here -->
    </div>
    <script>
        function button_clicked(buttonId) {
            var data = {
                "ao_thun": `[ao_thun]`,
                "chai_nuoc": `[chai_nuoc]`,
                "kep_toc": `[kep_toc]`,
                "op_dien_thoai": `[op_dien_thoai]`,
                "loa_bluetooth": `[loa_bluetooth]`
            };
            document.getElementById('display').innerText = data[buttonId];
        }
    </script>
</body>
</html>
"""

# Replace placeholders with actual DataFrame JSON data
html_content = html_template
for key, json_df in json_dataframes.items():
    html_content = html_content.replace(f'[{key}]', json_df)

# Generate buttons
buttons_html = ""
for key in new_comments_dict.keys():
    buttons_html += f'<button id="b_{key}" class="button-55" onclick="button_clicked(\'{key}\')">Show {key} comments</button><br><br>\n'

# Replace button placeholder in the template
html_content = html_content.replace("[button]", buttons_html)

# Write the final HTML content to a file
with open("example2.html", "w") as file:
    file.write(html_content)

# Optionally open the file in a web browser
os.startfile("example2.html")