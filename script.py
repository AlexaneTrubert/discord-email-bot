import discord
import imaplib
import email
from discord.ext import tasks
from bs4 import BeautifulSoup
from email.header import decode_header
import os
import tempfile

# TOKEN discord bot
TOKEN = 'TOKEN_DISCORD_BOT'

# Id for email
EMAIL = 'EMAIL'
PASSWORD = 'PASSWORD'

# Id for serveur, channel and role to ping
GUILD_ID =
CHANNEL_ID =
SUPPORT_ROLE_ID =

intents = discord.Intents.default()
intents.guilds = True
intents.messages = True
client = discord.Client(intents=intents)

# Function to download attachments
async def download_attachment(part):
    filename = part.get_filename()
    if bool(filename):
        filepath = os.path.join(tempfile.gettempdir(), filename)
        with open(filepath, 'wb') as f:
            f.write(part.get_payload(decode=True))
        return filepath
    return None

async def fetch_emails():
    mail = imaplib.IMAP4_SSL('imap-mail.outlook.com')
    mail.login(EMAIL, PASSWORD)
    mail.select('inbox')
    
    # Get all emails that are unread
    result, data = mail.search(None, 'UNSEEN')
    email_ids = data[0].split()

    if email_ids:
        print(f"{len(email_ids)} new emails found")

    # Loop through all the emails to get the content
    for email_id in email_ids:
        resp, raw_email = mail.fetch(email_id, '(BODY[])')
        raw_email = raw_email[0][1]
        msg = email.message_from_bytes(raw_email)
        
        # Get subject
        subject = decode_header(msg['subject'])[0][0]
        if isinstance(subject, bytes):
            # If it's a bytes, decode to str
            subject = subject.decode()
        
        content = ''
        attachments = []

        # Get the email content
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get('Content-Disposition'))

                if "attachment" in content_disposition:
                    # If it's an attachment, download it to the temp folder
                    filepath = await download_attachment(part)
                    if filepath is not None:
                        attachments.append(filepath)
                else:
                    # If it's not an attachment, add the content to the email content
                    raw_content = part.get_payload(decode=True)
                    if content_type == "text/html":
                        soup = BeautifulSoup(raw_content, "html.parser")
                        content = soup.get_text()
                    else:
                        content = raw_content
        else:
            raw_content = msg.get_payload(decode=True)
            soup = BeautifulSoup(raw_content, "html.parser")
            content = soup.get_text()

        # Write the email content to a discord message
        discord_message = f"<@&{SUPPORT_ROLE_ID}> **{subject}**\n\n{content}"

        # If the message is too long, split it into chunks
        discord_message_chunks = [discord_message[i:i+2000] for i in range(0, len(discord_message), 4000)]

        guild = client.get_guild(GUILD_ID)
        channel = guild.get_channel(CHANNEL_ID)

        # We send he message to discord
        for chunk in discord_message_chunks:
            await channel.send(chunk)
            print(f"We send new message on discord: {subject}")

        # Send the attachments
        files = []
        for i, filepath in enumerate(attachments):
            with open(filepath, 'rb') as f:
                files.append(discord.File(f, filename=os.path.basename(filepath)))
            # If we have 10 files or if it's the last file, send the files
            if (i + 1) % 10 == 0 or i == len(attachments) - 1:
                await channel.send(files=files)
                files = []

        try:
            mail.store(email_id, '+FLAGS', '\Seen')
        except Exception as e:
            print(f"Error when we pass email to seen: {e}")

    mail.close()
    mail.logout()

# When the bot is ready
@client.event
async def on_ready():
    print(f'Bot connected {client.user.name}')
    print('------')
    await fetch_emails()
    check_emails.start()

# Check emails every hour
@tasks.loop(hours=1)
async def check_emails():
    await fetch_emails()

# Launch the bot
client.run(TOKEN)
