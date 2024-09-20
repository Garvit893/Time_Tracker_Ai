import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
from email_validator import validate_email, EmailNotValidError
import os
from groq import Groq  


api_key = st.secrets["groq"]["api_key"]  
client = Groq(api_key=api_key)

MODEL = 'llama3-groq-70b-8192-tool-use-preview'

# Function to send an email
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

# Function to categorize reasons
def categorize_reason(reason):
    if "official" in reason.lower():
        return "Official"
    elif "emergency" in reason.lower():
        return "Emergency"
    elif "personal" in reason.lower():
        return "Due to Some Personal Work"
    elif "slept in the office" in reason.lower() or "emergency in office" in reason.lower():
        return "Shady"
    else:
        return "Not Genuine"

# Function to generate email body using Groq API
def generate_email_body(employee_name, reason, category):
    prompt = f"Dear {employee_name},\n\n"
    if category == "Shady":
        prompt += f"Your reported work hours are less than 48 hours this week due to the reason: '{reason}'. Please refrain from engaging in personal work during office hours.\n\nBest Regards,\nManagement"
    else:
        prompt += f"Your reported work hours are less than 48 hours this week due to the reason: '{reason}'. Please provide a valid reason.\n\nBest Regards,\nManagement"

   
    try:
        response = client.chat(
            model=MODEL,
            messages=[{"role": "user", "content": prompt}],
        )
        
        if response and 'choices' in response and len(response['choices']) > 0:
            return response['choices'][0]['message']['content']
        else:
            return "Error generating email content."
    except Exception as e:
        return f"An error occurred while generating content: {str(e)}"

# Streamlit UI
st.title('Work Hour Tracker')

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Assuming the Excel file has 'Employee Name', 'Email', 'Work Hours', 'Reason' columns
    defaulters = df[df['Work Hours'] < 48]
    
    approved_reasons = []
    not_genuine_reasons = []
    shady_reasons = []
    email_list = []

    for _, row in defaulters.iterrows():
        employee_name = row['Employee Name']
        email = row['Email']
        reason = row['Reason']
        
        category = categorize_reason(reason)
        if category in ["Official", "Emergency", "Due to Some Personal Work"]:
            approved_reasons.append((employee_name, email, reason, category))
        elif category == "Shady":
            shady_reasons.append((employee_name, email, reason))
            email_list.append(email)
            subject = f"Attendance Alert for {employee_name}"
            body = generate_email_body(employee_name, reason, category)
            send_email(email, subject, body)
        else:
            not_genuine_reasons.append((employee_name, email, reason, category))
            email_list.append(email)
            subject = f"Attendance Alert for {employee_name}"
            body = generate_email_body(employee_name, reason, category)
            send_email(email, subject, body)

    # Display results
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

    # Save results to Excel
    output_df = pd.concat([approved_df, not_genuine_df, shady_df], axis=0)
    output_file = 'defaulter_results.xlsx'
    output_df.to_excel(output_file, index=False)
    st.success(f"Results saved to {output_file}")
