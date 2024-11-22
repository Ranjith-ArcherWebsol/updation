import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import ssl

# Streamlit UI for file upload
st.title("Automated updation Email Notification with File Upload")

uploaded_file = st.file_uploader("Upload Excel File", type="xlsx")

if uploaded_file:
    # Read the uploaded Excel file
    client_data = pd.read_excel(uploaded_file)
    st.write("Uploaded Data:", client_data)

    # Email fields
    sender_email = "archerwebsol@gmail.com"
    smtp_server = "smtp.gmail.com"
    port = 465  # For SSL
    password = st.text_input("Enter your email password", type="password")  # Keep this secure

    # Extracting URL for the subject (you can choose a specific row or value)
    if 'url' in client_data.columns:
        url = client_data.iloc[0]['url']
    else:
        url = "No URL"  # Default value in case 'url' column doesn't exist

    # Email subject and HTML content
    subject = f"REQUEST FOR UPDATION - {url}"

    # HTML formatted email content
    html_body = """\
    Dear Sir / Madam,<br><br>

    Greetings from Archer Websol...<br><br>

    Please send us any content and images for update on your website.<br><br>

    Once received your mail, we will update as soon as possible and let you know immediately.<br><br>

    Awaiting your reply.
    """

    # Display the HTML preview in Streamlit
    st.write("Email Preview:")
    st.markdown(html_body, unsafe_allow_html=True)

    # Send email button
    if st.button("Send Email"):
        # Iterate over each email in the 'email' column of the uploaded data
        for email in client_data['email']:
            # Set up the email
            message = MIMEMultipart("alternative")
            message["From"] = sender_email
            message["To"] = email
            message["Subject"] = subject
            message.attach(MIMEText(html_body, "html"))

            # Send the email
            try:
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
                    server.login(sender_email, password)
                    server.sendmail(sender_email, email, message.as_string())
                st.success(f"Email sent successfully to {email}!")
            except Exception as e:
                st.error(f"Error sending email to {email}: {e}")
