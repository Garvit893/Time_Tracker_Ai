import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
from email_validator import validate_email, EmailNotValidError
from groq import Groq

# Set up Streamlit secrets
api_key = "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"

# Set up Groq client
client = Groq(api_key=api_key)
model = 'llama3-70b-8192'

# Function to send email
def send_email(recipient, subject, body):
    email_address = 'garvit@marketingmindz.in'
    email_password = 'GFsJ271b'
    
    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = email_address
        msg['To'] = recipient
        msg.set_content(body)

        with smtplib.SMTP('smtp.marketingmindz.in', 587) as smtp:
            smtp.starttls()
            smtp.login(email_address, email_password)
            smtp.send_message(msg)
        return True
    except Exception as e:
        st.error(f"Error sending email: {str(e)}")
        return False

# Function to process reason and generate email
def process_reason_and_generate_email(employee_name, reason):
    try:
        prompt = (
    f"Categorize the following reason and generate an email body:\n\n"
    f"Employee Name: {employee_name}\n"
    f"Reason: {reason}\n\n"
    f"Categories: Official, Emergency, Personal, Shady.\n"
    f"Greet {employee_name} appropriately and use a tone fitting for a manager and the report is for the whole week.\n"
    f"The name in the sign-off/closing should be 'HR Team'.\n"
    f"Please respond in the following format:\n"
    f"Category: <category>\n"
    f"Email Body: <email body>"
)

        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}]
        )

        if response:
            content = response.choices[0].message.content.strip()
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
        st.error(f"Error processing reason: {str(e)}")
        return "Error", "An error occurred while generating content."

# Streamlit app
st.title('Work Hour Tracker')

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        st.stop()

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
            if send_email(email, subject, body):
                email_list.append(email)
        elif category == "Shady":
            shady_reasons.append((employee_name, email, reason))
            if send_email(email, subject, body):
                email_list.append(email)
        else:
            not_genuine_reasons.append((employee_name, email, reason, category))
            if send_email(email, subject, body):
                email_list.append(email)

    st.subheader("Approved Reasons")
    approved_df = pd.DataFrame(approved_reasons, columns=["Employee Name", "Email", "Reason", "Category"])
    st.table(approved_df)

    st.subheader("Not Genuine Reasons")
    not_genuine_df = pd.DataFrame(not_genuine_reasons, columns=["Employee Name", "Email", "Reason", "Category"])
    st.table(not_genuine_df)

    st.subheader("Shady Reasons")
    shady_df = pd .DataFrame(shady_reasons, columns=["Employee Name", "Email", "Reason"])
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
