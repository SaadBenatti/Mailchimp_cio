""" 
All you need to do is edit the fields at the bottom with your 
audience id and api key and it should work, I would do a test run first 
because I think it has problem with images
"""


import pandas as pd
from mailchimp_marketing import Client
from mailchimp_marketing.api_client import ApiClientError

# Read the Excel file
def read_excel(file_path):
    df = pd.read_excel(file_path)
    return df

# Generate HTML content for the newsletter
def generate_html(df):
    html_content = "<html><body>"
    
    for index, row in df.iterrows():
        title = row['Title']
        image = row['Image']
        description = row['Description']
        
        html_content += f"""
        <div style="margin-bottom: 20px;">
            <h2>{title}</h2>
            <img src="{image}" style="width:100%;max-width:600px;" />
            <p>{description}</p>
        </div>
        """
        
    html_content += "</body></html>"
    return html_content

# Create a campaign in Mailchimp
def create_campaign(mailchimp, audience_id, subject, from_email, from_name, content):
    try:
        response = mailchimp.campaigns.create({
            "type": "regular",
            "recipients": {"list_id": audience_id},
            "settings": {
                "subject_line": subject,
                "reply_to": from_email,
                "from_name": from_name,
            },
        })
        campaign_id = response['id']

        # Set the content of the campaign
        mailchimp.campaigns.set_content(campaign_id, {"html": content})

        # Send the campaign
        mailchimp.campaigns.send(campaign_id)
        print(f"Campaign sent successfully! ID: {campaign_id}")
    except ApiClientError as error:
        print(f"An error occurred: {error.text}")

# Main function
def main(file_path, api_key, audience_id, subject, from_email, from_name):
    # Initialize Mailchimp client
    mailchimp = Client()
    mailchimp.set_config({"api_key": api_key, "server": api_key.split('-')[-1]})

    # Read Excel data
    df = read_excel(file_path)

    # Generate HTML for the newsletter
    html_content = generate_html(df)

    # Create and send Mailchimp campaign
    create_campaign(mailchimp, audience_id, subject, from_email, from_name, html_content)
    
#EDIT HERE BELOW

if __name__ == "__main__":
    file_path = "newsletter_data.xlsx"  # Path to your Excel file, You can change the name
    api_key = "b29e0d807288c91da394f7e7c568b072-us13" #TEST API KEY, REMOVE
    audience_id = "5e858125b0" #TEST AUDIENCE ID, REMOVE
    subject = "Your Newsletter Subject" 
    from_email = "Youremail@email.com"
    from_name = "Your Name"

    main(file_path, api_key, audience_id, subject, from_email, from_name)
