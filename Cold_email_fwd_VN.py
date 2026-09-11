import pandas as pd
import datetime
import pyperclip

# === PART 1: Read the Excel file and filter the rows to send emails today. ===
file_path = r"D:\**Your CEA_sample_input file path" ## Please paste the path to your sample file 'CEA_sample_input' here.
sheet_name = "Cold_email_FWD_VN"

# Reading data from the Excel file.
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Get the date variable = today
today = datetime.datetime.today().date()

# Follow-up columns
followup_columns = [
    "Follow-up 1 Due",
    "Follow-up 2 Due",
    "Follow-up 3 Due",
    "Follow-up 4 Due",
]

print("Các dòng cần gửi follow-up hôm nay (dựa theo dòng thực tế trong Excel):")
for idx, row in df.iterrows():
    for i, col in enumerate(followup_columns):
        due_date = row[col]
        if pd.notna(due_date):
            if isinstance(due_date, pd.Timestamp):
                due_date = due_date.date()
            if due_date == today:
                fl = f"FL{i+1}"
                excel_row = idx + 2  # Add 2 to match the actual row in Excel.
                print(f"{excel_row}:{fl}")

# === PART 2: Manually inputting and automatically copying email content ===
print("\n---TẠO NỘI DUNG EMAIL THEO YÊU CẦU THỦ CÔNG---")
print("Nhập cú pháp theo dạng: số dòng Excel và FL tương ứng (VD: 4 FL1)")
print("Nhập 'x' để kết thúc.")

# Email template – add all variables
email_templates = {
    "FL1": """Dear {contact_person}, 

Chúc {xung_ho} một ngày nhiều niềm vui!
Em hy vọng có cơ hội được đồng hành cùng {company_name} trong thời gian tới.

Đây là follow up email thứ nhất.

""", #You can adjust your own email template here! I suggest to add a keyword and use the autotext function in Outlook, where it's mỏe easy to use collor text, images, etc.
    "FL2": """Hi {contact_person} ơi,

Hôm trước em có giới thiệu sơ về dịch vụ mà bên em cung cấp.
Tuy nhiên có một số chi tiết chưa thể trình bày trong một email được:
...
Hy vọng những ưu điểm này hỗ trợ được cho {company_name}.
Cảm ơn {xung_ho} đã đọc email

""",
    "FL3": """
Dear {contact_person},

Hôm trước em có giới thiệu sơ về dịch vụ mà bên em cung cấp.
Tuy nhiên có một số chi tiết chưa thể trình bày trong một email được, hy vọng những ưu điểm này hỗ trợ được cho {company_name}.
...
Cảm ơn {xung_ho} đã đọc email

 """,
    "FL4": """Hi {contact_person}

Em xin phép gửi {xung_ho} nốt một email giới thiệu nữa {xung_ho} nha.
Bữa giờ em mail {xung_ho} quá trời mà hong thấy {xung_ho} phản hồi! Sợ e mail chưa đúng nhân sự phụ trách hoặc có thể thời điểm này mình chưa có nhu cầu ^ ^

Em có đính kèm lại mấy email trước cùng file giá để khi cần {xung_ho} cũng tiện search lại thông tin bên em – Topaz Marine!
{xung_ho} nhận được email cho em xin một phản hồi để em biết mình gửi đúng PIC nha.

Chúc {xung_ho} ngày mới thuận lợi. Em vẫn luôn ready để làm việc cùng {xung_ho}!

""",
}

while True:
    user_input = input("Nhập dòng và FL: ")
    if user_input.lower() == "x":
        break
    try:
        excel_row_str, fl = user_input.strip().split()
        excel_row = int(excel_row_str)
        fl = fl.upper()
        index = excel_row - 2  # Trừ lại 2 để khớp với chỉ số trong dataframe

        if index < 0 or index >= len(df):
            print("Số dòng Excel không hợp lệ.")
            continue
        if fl not in email_templates:
            print("Chỉ chấp nhận FL1 đến FL4.")
            continue

        company = df.at[index, "Company Name"]
        contact_person = df.at[index, "Contact Person"]
        xung_ho = df.at[index, "Xưng hô"]
        email = df.at[index, "Email"]

        # Format email content with variables
        message = email_templates[fl].format(
            xung_ho=xung_ho, contact_person=contact_person, company_name=company
        )

        email_text = f"""To: {email}
Subject: Follow-up from [Your Company] – {company}

{message}"""

        pyperclip.copy(email_text)

        print("\n--- EMAIL ĐÃ COPY VÀO CLIPBOARD ---")
        print(email_text)
        print("Đã sao chép. Bạn chỉ cần Ctrl+V vào mail.")
        print("----------------------\n")

    except Exception as e:
        print(f"Lỗi nhập: {e}. Vui lòng nhập lại theo cú pháp.")

