import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
from email_validator import validate_email, EmailNotValidError
from groq import Groq

api_key = st.secrets["groq"]["api_key"]
client = Groq(api_key=api_key)
MODEL = 'llama3-groq-70b-8192'

def send_email(recipient, subject, body):
    EMAIL_ADDRESS = st.secrets["general"]["email_address"]
    EMAIL_PASSWORD = st.secrets["general"]["email_password"] 

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = recipient
    msg.set_content(body)

    with smtplib.SMTP('smtp.marketingmindz.in', 587) as smtp:
        smtp.starttls()
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)

def process_reason_and_generate_email(employee_name, reason):
    prompt = (
        f"Categorize the following reason and generate an email body:\n\n"
        f"Employee Name: {employee_name}\n"
        f"Reason: {reason}\n\n"
        f"Categories: Official, Emergency, Personal, Shady.\n"
        f"Please respond in the following format:\n"
        f"Category: <category>\n"
        f"Email Body: <email body>"
    )
    
    try:
        response = client.chat_complete(  
            model=MODEL,
            messages=[{"role": "user", "content": prompt}]
        )

        print("Prompt sent to AI:", prompt)  
        print("AI Response:", response)  

        if response and 'choices' in response and len(response['choices']) > 0:
            content = response['choices'][0]['message']['content'].strip()
            print("Raw AI Content:", content)  
          
            if "Category:" in content and "Email Body:" in content:
                category_line, body_line = content.split("Email Body:")
                category = category_line.split("Category:")[1].strip()
                body = body_line.strip()
                return category, body
            else:
                return "Unknown", "Could not parse the response correctly."
        else:
            return "Error", "Unexpected response structure."
    except Exception as e:
        return "Error", f"An error occurred while generating content: {str(e)}"

st.title('Work Hour Tracker')

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)

    defaulters = df[df['Work Hours'] < 48]
    
    approved_reasons = []
    not_genuine_reasons = []
    shady_reasons = []
    email_list = []

    for _, row in defaulters.iterrows():
        employee_name = row['Employee Name']
        email = row['Email']
        reason = row['Reason']

        if pd.isna(email) or not isinstance(email, str) or "@" not in email:
            st.warning(f"Invalid email for {employee_name}. Skipping.")
            continue
        
        try:
            validate_email(email)
        except EmailNotValidError:
            st.warning(f"Invalid email for {employee_name}. Skipping.")
            continue
        
        category, body = process_reason_and_generate_email(employee_name, reason)
        subject = f"Attendance Alert for {employee_name}"

        if category in ["Official", "Emergency", "Personal"]:
            approved_reasons.append((employee_name, email, reason, category))
            send_email(email, subject, body)
        elif category == "Shady":
            shady_reasons.append((employee_name, email, reason))
            email_list.append(email)
            send_email(email, subject, body)
        else:
            not_genuine_reasons.append((employee_name, email, reason, category))
            email_list.append(email)
            send_email(email, subject, body)

    st.subheader("Approved Reasons")
    approved_df = pd.DataFrame(approved_reasons, columns=["Employee Name", "Email", "Reason", "Category"])
    st.table(approved_df)

    st.subheader("Not Genuine Reasons")
    not_genuine_df = pd.DataFrame(not_genuine_reasons, columns=["Employee Name", "Email", "Reason", "Category"])
    st.table(not_genuine_df)

    st.subheader("Shady Reasons")
    shady_df = pd.DataFrame(shady_reasons, columns=["Employee Name", "Email", "Reason"])
    st.table(shady_df)

    if email_list:
        st.success("Emails sent to the following employees:")
        st.write(", ".join(email_list))
    else:
        st.info("No emails sent; all reasons were valid.")

    output_df = pd.concat([approved_df, not_genuine_df, shady_df], axis=0)
    output_file = 'defaulter_results.xlsx'
    output_df.to_excel(output_file, index=False)

    with open(output_file, "rb") as f:
        st.download_button(
            label="Download Results",
            data=f,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    st.success(f"Results saved to {output_file}")
