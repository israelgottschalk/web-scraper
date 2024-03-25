import win32com.client as win32

def send_outlook_email(receiver_email, subject, body, attachment_path=None):
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)  # Represents an Outlook mail item

    # Fill in email details
    mail.To = receiver_email
    mail.Subject = subject
    mail.Body = body

    # Attach file if provided
    if attachment_path:
        attachment = attachment_path
        mail.Attachments.Add(attachment)

    # Display or send the email
    mail.Display()  # Use mail.Send() to send without displaying
      # Send the email
    mail.Send()  # Use mail.Display() to open in Outlook without sending

# Usage example
if __name__ == "__main__":
    receiver_email = 'PSRIntelligence@psr.org.uk'
    subject = 'Scraped news'
    body = 'Here are the scraped news from The Paypers and Finextra.'
    attachment_path = 'C:/Users/IGottschalk/Downloads/women_men.pdf'  # Optional

    send_outlook_email(receiver_email, subject, body, attachment_path)
