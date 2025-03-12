# for data processing
import pandas as pd

# visualisation
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px

# sending the report
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# handing files and potential errors
import sys
import os

# debugging (check current directory)
print(f"Current Directory: {os.getcwd()}")

file_path = "Coffee Shop Sales.xlsx"

# checking if this file exists
if not os.path.exists(file_path):
    print(f"Error: File '{file_path}' not found.")
    sys.exit(1)
try:
    df = pd.read_excel(file_path)
    print("Excel file loaded successfully.")
except Exception as e:
    print(f"Something went wrong while loading the file: {e}")
    sys.exit(1)

# checking if the data column is correct
if "transaction_date" not in df.columns:
    print("Error: 'transaction_date' column is missing.")
    sys.exit(1)
try:
    df["transaction_date"] = pd.to_datetime(df["transaction_date"])
except Exception as e:
    print(f"Error during date conversion: {e}")
    sys.exit(1)

# sorting data by year
df_year_2023 = df[df["transaction_date"].dt.year == 2023]

# checking if column 'product_category' exists at all
if "product_category" not in df.columns:
    print("Error: 'product_category' column is missing.")
    sys.exit(1)

# sorting the category
df_coffee_2023 = df_year_2023[df_year_2023["product_category"] == "Coffee"].copy()

# adding the 'month' column
df_coffee_2023.loc[:, 'month'] = df_coffee_2023['transaction_date'].dt.month

# grouping by store location and summing up sales
df_top_shops = df_coffee_2023.groupby("store_location")["transaction_qty"].sum()

# getting top 5 shops
df_top_5_shops = df_top_shops.sort_values(ascending=False).head(5)

print("Top 5 sklepów z najwyższą sprzedażą kawy w 2023 roku:")
print(df_top_5_shops)

# visualisation
plt.figure(figsize=(10, 5))
sns.barplot(x=df_top_5_shops.values, y=df_top_5_shops.index, hue=df_top_5_shops.index, palette="viridis", legend=False)
plt.xlabel("Liczba sprzedanych jednostek kawy")
plt.ylabel("Lokalizacja sklepu")
plt.title("Top 5 sklepów z najwyższą sprzedażą kawy w 2023 roku")
plt.savefig("top_5_shops.png")
plt.show(block=True)  # Ensure it does not close immediately

# monthly coffee sales analysis
monthly_sales = df_coffee_2023.groupby('month')['transaction_qty'].sum()

# visualisation (by month)
plt.figure(figsize=(10, 5))
sns.lineplot(x=monthly_sales.index, y=monthly_sales.values, marker="o")
plt.xlabel("Miesiąc")
plt.ylabel("Liczba sprzedanych jednostek kawy")
plt.title("Miesięczna sprzedaż kawy w 2023 roku")
plt.savefig("monthly_sales.png")
plt.show(block=True)

# Interactive chart with Plotly
fig = px.line(monthly_sales, x=monthly_sales.index, y=monthly_sales.values,
              title="Miesięczna sprzedaż kawy w 2023 roku")

# Debugging: Save instead of show
fig.write_html("plotly_chart.html")
print("Plotly chart saved as 'plotly_chart.html'. Open it in a browser.")

# Generating the report
report = f"""
Coffee Sales Report 2023

Top 5 Stores with the Highest Coffee Sales
{df_top_5_shops.to_markdown()}

Monthly Coffee Sales Trend
"""

# saving the report to a Markdown file
report_filename = "coffee_sales_report.md"
with open(report_filename, "w", encoding="utf-8") as f:
    f.write(report)

print(f"Report saved as '{report_filename}'.")


# function to send the report
def email_sending(subject, body, to_email, attachment_path):
    from_email = "webapplicationinteractiontest@gmail.com"
    from_password = "ckzy nefk rqai wbhw"  # Use App Password instead!

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Attach file
    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(attachment_path)}"')
            msg.attach(part)
    else:
        print(f"Warning: File '{attachment_path}' not found. Email will be sent without an attachment.")

    # Send email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, from_password)
        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()
        print("Email sent successfully.")
    except Exception as excp:
        print(f"Error sending email: {excp}")


# Sending coffee sales report
email_sending(
    subject="Coffee Sales Report 2023",
    body="Attached is the coffee sales report for the year 2023.",
    to_email="oyleksa@gmail.com",
    attachment_path=report_filename
)

print("Script execution completed.")
