import os
import re
import time
import pandas as pd
import win32com.client
import win32gui
import win32con
import threading
import urllib.parse
from datetime import datetime, timedelta
import logging
from collections import defaultdict
import uuid

print("="*80)
print("INTELLIGENT MEETING SCHEDULER BOT - FIXED VERSION")
print("="*80)


# ════════════════════════════════════════════════════════════════════════════
# ✅ AUTO-CLICK FORTRA POPUP FUNCTIONS
# ════════════════════════════════════════════════════════════════════════════

def _find_classification_dialog():
    """Find FORTRA popup window"""
    hwnds = []
    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if not title:
                return
            title_lower = title.lower()
            if title_lower.startswith("dcs -") or any(
                k in title_lower for k in ("fortra", "classif")
            ):
                hwnds.append(hwnd)
    win32gui.EnumWindows(enum_handler, None)
    return hwnds[0] if hwnds else None


def _click_ok_via_uia(hwnd):
    """Click OK using UI Automation"""
    try:
        from pywinauto import Application
        app = Application(backend="uia").connect(handle=hwnd)
        win = app.window(handle=hwnd)
        win.set_focus()
        ok_btn = win.child_window(title="OK", control_type="Button")
        ok_btn.click_input()
        return True
    except Exception:
        return False


def _click_dialog_button(hwnd):
    """Click dialog button with fallback"""
    if _click_ok_via_uia(hwnd):
        return
    clicked = []
    def enum_child(child_hwnd, _):
        if win32gui.GetClassName(child_hwnd) == "Button":
            label = win32gui.GetWindowText(child_hwnd).strip().lower()
            if label in ("ok", "yes", "send", "confirm", "close", "continue"):
                win32gui.PostMessage(child_hwnd, win32con.BM_CLICK, 0, 0)
                clicked.append(child_hwnd)
    win32gui.EnumChildWindows(hwnd, enum_child, None)
    if not clicked:
        try:
            win32gui.SetForegroundWindow(hwnd)
            win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
            win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
        except Exception:
            pass


def _auto_dismiss_classification_dialog(stop_event, poll_interval=0.25):
    """Background thread monitoring for FORTRA popup every 0.25 seconds"""
    while not stop_event.is_set():
        hwnd = _find_classification_dialog()
        if hwnd:
            _click_dialog_button(hwnd)
        stop_event.wait(poll_interval)


def send_with_auto_classify(send_callable):
    """Send email with auto-popup dismissal - NO MANUAL CLICKS NEEDED"""
    stop_event = threading.Event()
    watcher = threading.Thread(
        target=_auto_dismiss_classification_dialog, args=(stop_event,), daemon=True
    )
    watcher.start()
    try:
        send_callable()
    finally:
        stop_event.set()

# ════════════════════════════════════════════════════════════════════════════

# Configuration
DATA_DIR = os.getenv("KFSHRC_DATA_DIR", os.path.expanduser("~/KFSHRC_data"))
os.makedirs(DATA_DIR, exist_ok=True)

LOG_FILE = os.path.join(DATA_DIR, "meeting_log.csv")
RESPONSES_FILE = os.path.join(DATA_DIR, "meeting_responses.csv")
TIMELINE_FILE = os.path.join(DATA_DIR, "meeting_timeline.csv")
METADATA_FILE = os.path.join(DATA_DIR, "meeting_metadata.csv")  # ← NEW: Store original attendees
ERROR_LOG = os.path.join(DATA_DIR, "error_log.txt")

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(ERROR_LOG),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

BOT_SUBJECT_KEYWORD = "meeting bot"
CHECK_INTERVAL = 10

OL_FOLDER_INBOX = 6
OL_APPOINTMENT_ITEM = 1
OL_MEETING = 1
OL_REQUIRED = 1

WORK_START_HOUR = 9
WORK_END_HOUR = 17
DEFAULT_DURATION_MIN = 60

WEEKEND_DAYS = {4, 5}

EMAIL_REGEX = r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9.-]+'
# Matches Outlook's plain-text rendering of a hyperlink whose display text
# differs from its href, e.g. "shown@x.com <mailto:actual@x.com>". The
# hidden mailto: target is often a stale/wrong autocomplete match, so we
# keep only the text the sender actually typed and drop the "<mailto:...>" part.
MAILTO_DISPLAY_MISMATCH_REGEX = (
    r'([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9.-]+)'
    r'\s*<mailto:[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9.-]+>'
)

# Addresses that must never be used as meeting attendees/survey recipients,
# regardless of how they show up in a request body (e.g. bad contact
# resolution on the sender's side). Add lowercase addresses here.
BLACKLISTED_EMAILS = {
    'halhazmi@kfshrc.edu.sa',  # ← ONLY THIS EMAIL IS BLACKLISTED
}
DATETIME_REGEX = r'(\d{2}-\d{2}-\d{4})\s+(\d{2}:\d{2})'

logger.info(f"DATA_DIR: {DATA_DIR}")
logger.info(f"RESPONSES_FILE: {RESPONSES_FILE}")
logger.info("Connecting to Outlook...")


# ════════════════════════════════════════════════════════════════════════════
# ✅ NEW: TEAMS + CALENDAR URL GENERATION FUNCTIONS
# ════════════════════════════════════════════════════════════════════════════

def generate_teams_meeting_url(meeting_id, topic):
    """Generate Teams meeting URL - no auth needed to join"""
    base_url = "https://teams.microsoft.com/l/meetup-join"
    teams_url = f"{base_url}/{meeting_id}?context=%7B%7D"
    if not meeting_id:
        teams_url = "https://www.microsoft.com/en-us/microsoft-teams/download-app"
    return teams_url


def generate_google_calendar_url(topic, start_datetime, duration_minutes, attendees):
    """Generate Google Calendar event URL"""
    start_time = start_datetime.strftime('%Y%m%dT%H%M%S')
    end_time = (start_datetime + timedelta(minutes=duration_minutes)).strftime('%Y%m%dT%H%M%S')
    title = f"Meeting: {topic}"
    description = f"Attendees: {', '.join(attendees)}"
    params = {
        'action': 'TEMPLATE',
        'text': title,
        'dates': f"{start_time}/{end_time}",
        'details': description,
    }
    google_cal_url = "https://calendar.google.com/calendar/render?" + urllib.parse.urlencode(params)
    return google_cal_url


def generate_outlook_calendar_url(topic, start_datetime, duration_minutes, attendees):
    """Generate Outlook calendar URL"""
    title = f"Meeting: {topic}"
    attendees_str = "; ".join(attendees)
    params = {
        'subject': title,
        'startdt': start_datetime.isoformat(),
        'enddt': (start_datetime + timedelta(minutes=duration_minutes)).isoformat(),
        'body': f"Attendees: {attendees_str}",
        'location': 'Teams Meeting',
    }
    outlook_url = "https://outlook.live.com/calendar/0/deeplink/compose?" + urllib.parse.urlencode(params)
    return outlook_url


# ════════════════════════════════════════════════════════════════════════════

class MeetingBot:
    def __init__(self):
        self.processed_requests = set()
        self.processed_responses = set()
        self.connect_outlook()

    def connect_outlook(self):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.inbox = self.namespace.GetDefaultFolder(OL_FOLDER_INBOX)
            logger.info("Outlook connected!")
        except Exception as e:
            logger.error(f"Outlook connection failed: {e}")
            raise

    def get_sender_smtp_address(self, msg):
        try:
            PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
            smtp = msg.PropertyAccessor.GetProperty(PR_SENDER_SMTP_ADDRESS)
            if smtp:
                return smtp
        except:
            pass
        try:
            sender = msg.Sender
            if sender and sender.AddressEntry.Type == "EX":
                exchange_user = sender.AddressEntry.GetExchangeUser()
                if exchange_user and exchange_user.PrimarySmtpAddress:
                    return exchange_user.PrimarySmtpAddress
        except:
            pass
        return msg.SenderEmailAddress

    def parse_attendees(self, body):
        # Try to extract attendees from "Attendees: " line
        # Handle both single-line and multi-line attendees
        match = re.search(r'attendees\s*:\s*(.+?)(?=\n(?:topic|duration|deadline|$))', body, re.I | re.DOTALL)
        if match:
            attendees_text = match.group(1)
            # Outlook renders "Display Text <mailto:actual@address>" for links whose
            # display text doesn't match their href. Keep only the display text
            # (what the sender actually typed) and drop the "<mailto:...>" part,
            # otherwise one attendee gets split into two - see meeting 25280690.
            attendees_text = re.sub(MAILTO_DISPLAY_MISMATCH_REGEX, r'\1', attendees_text)
            # Extract all emails from the attendees section
            emails = re.findall(EMAIL_REGEX, attendees_text)
        else:
            emails = []

        logger.info(f"Raw attendees extracted: {emails}")
        
        # Remove duplicates and check blacklist - STRICT FILTERING!
        seen = set()
        result = []
        blacklist_count = 0
        
        for email in emails:
            key = email.lower()
            if key in seen:
                continue
            if key in BLACKLISTED_EMAILS:
                logger.info(f"[BLACKLIST BLOCKING] {email} is BLACKLISTED - WILL NOT SEND SURVEY")
                blacklist_count += 1
                continue
            seen.add(key)
            result.append(email)
        
        if blacklist_count > 0:
            logger.info(f"[BLACKLIST SUMMARY] Blocked {blacklist_count} blacklisted email(s)")
        logger.info(f"[FINAL ATTENDEES] Sending surveys to: {result}")
        return result

    def parse_duration(self, body):
        match = re.search(r'duration\s*:\s*(\d+)', body, re.I)
        return int(match.group(1)) if match else DEFAULT_DURATION_MIN

    def parse_topic(self, body, fallback_subject):
        match = re.search(r'topic\s*:\s*(.+)', body, re.I)
        if match:
            return match.group(1).strip()
        return re.sub(BOT_SUBJECT_KEYWORD, '', fallback_subject, flags=re.I).strip(' -:') or "Meeting"

    def parse_deadline(self, body):
        match = re.search(r'deadline\s*:\s*(\d{2}-\d{2}-\d{4})\s+(\d{2}:\d{2})', body, re.I)
        if match:
            try:
                return datetime.strptime(f"{match.group(1)} {match.group(2)}", '%d-%m-%Y %H:%M')
            except:
                pass
        return datetime.now() + timedelta(minutes=12)

    def extract_all_datetimes_from_body(self, body):
        """Extract ALL date/time combinations from body"""
        matches = re.findall(DATETIME_REGEX, body)
        datetimes = []
        
        for date_str, time_str in matches:
            try:
                dt = datetime.strptime(f"{date_str} {time_str}", '%d-%m-%Y %H:%M')
                if dt not in datetimes:  # Avoid duplicates
                    datetimes.append(dt)
            except:
                pass
        
        return datetimes

    def extract_meeting_request(self, body, subject, organizer_email):
        # Organizer is surveyed like any other attendee, whether or not they
        # typed their own address into the "Attendees:" line.
        attendees = self.parse_attendees(body)
        
        # IMPORTANT: Check if organizer is blacklisted before adding them!
        if organizer_email:
            organizer_lower = organizer_email.lower()
            if organizer_lower in BLACKLISTED_EMAILS:
                logger.info(f"[BLACKLIST] Organizer {organizer_email} is blacklisted - EXCLUDED from survey!")
            elif organizer_lower not in [a.lower() for a in attendees]:
                attendees.append(organizer_email)
        
        return {
            "attendees": attendees,
            "duration": self.parse_duration(body),
            "topic": self.parse_topic(body, subject),
            "deadline": self.parse_deadline(body),
        }

    def send_survey_email_arabic(self, attendee_email, meeting_id, topic, organizer, deadline):
        """Send survey in English - simple, no unsupported properties"""
        try:
            mail = self.outlook.CreateItem(0)
            mail.To = attendee_email
            # process_responses() matches replies by finding "ID: <id>)" in the
            # subject - it must be present here or replies are silently ignored.
            mail.Subject = f"{topic} (ID: {meeting_id})"
            # DO NOT set Categories - causes Outlook error
            # DO NOT set MarkAsTask - causes Outlook error
            # Only use supported properties
            mail.Body = f"""Hello,

You have been invited to the following meeting:

Topic:      {topic}
Organizer:  {organizer}
Meeting ID: {meeting_id}

Please reply to this email with the times you are available:
DD-MM-YYYY HH:MM

Example reply:
02-08-2026 09:00
03-08-2026 14:30

Thank you!
"""
            # AUTO-CLICK: Use send_with_auto_classify to auto-dismiss FORTRA popup
            send_with_auto_classify(mail.Send)
            logger.info(f"[AUTO-CLICK] Survey sent to {attendee_email} (FORTRA popup auto-closed)")
            self.log_timeline(meeting_id, "survey_sent", f"Survey sent to {attendee_email}")
            
        except Exception as e:
            logger.error(f"Failed to send survey to {attendee_email}: {e}")

    def get_unread_messages(self):
        try:
            messages = self.inbox.Items
            messages.Sort("[ReceivedTime]", True)

            result = []
            for i in range(min(100, messages.Count)):
                try:
                    msg = messages.Item(i + 1)
                    subject = msg.Subject or ""
                    is_unread = msg.UnRead
                    
                    # Check if it's a bot message (original survey) OR a reply with (ID:)
                    is_bot_message = BOT_SUBJECT_KEYWORD.lower() in subject.lower()
                    is_bot_reply = "(ID:" in subject  # Replies have (ID: meeting_id)
                    
                    if (is_bot_message or is_bot_reply) and is_unread:
                        msg_id = msg.EntryID
                        if msg_id not in self.processed_responses:
                            result.append(msg)
                except:
                    pass
            return result
        except Exception as e:
            logger.error(f"Failed to fetch messages: {e}")
            return []

    def save_all_responses(self, meeting_id, organizer, attendee, proposed_datetimes):
        """Save each proposed datetime as separate row in CSV"""
        try:
            response_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            for proposed_datetime in proposed_datetimes:
                response_entry = {
                    'Meeting ID': meeting_id,
                    'Organizer': organizer,
                    'Attendee': attendee,
                    'Proposed Date': proposed_datetime.strftime('%d-%m-%Y'),
                    'Proposed Time': proposed_datetime.strftime('%H:%M'),
                    'Response Time': response_time,
                    'Status': 'responded'
                }

                if os.path.exists(RESPONSES_FILE):
                    df = pd.read_csv(RESPONSES_FILE)
                    df = pd.concat([df, pd.DataFrame([response_entry])], ignore_index=True)
                else:
                    df = pd.DataFrame([response_entry])

                df.to_csv(RESPONSES_FILE, index=False, encoding='utf-8')
            
            logger.info(f"Saved {len(proposed_datetimes)} time options from {attendee}")
        except Exception as e:
            logger.error(f"Failed to save responses: {e}")

    def log_timeline(self, meeting_id, stage, description):
        try:
            timeline_entry = {
                'Meeting ID': meeting_id,
                'Stage': stage,
                'Description': description,
                'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }

            if os.path.exists(TIMELINE_FILE):
                df = pd.read_csv(TIMELINE_FILE)
                df = pd.concat([df, pd.DataFrame([timeline_entry])], ignore_index=True)
            else:
                df = pd.DataFrame([timeline_entry])

            df.to_csv(TIMELINE_FILE, index=False, encoding='utf-8')
        except Exception as e:
            logger.error(f"Failed to log timeline: {e}")

    def save_meeting_metadata(self, meeting_id, topic, organizer, attendees, deadline):
        """Save original meeting metadata (attendees list) for later retrieval"""
        try:
            metadata_entry = {
                'Meeting ID': meeting_id,
                'Topic': topic,
                'Organizer': organizer,
                'Original Attendees': ','.join(attendees),  # Store as comma-separated list
                'Deadline': deadline.strftime('%Y-%m-%d %H:%M:%S') if deadline else '',
                'Created At': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }

            if os.path.exists(METADATA_FILE):
                df = pd.read_csv(METADATA_FILE)
                df = pd.concat([df, pd.DataFrame([metadata_entry])], ignore_index=True)
            else:
                df = pd.DataFrame([metadata_entry])

            df.to_csv(METADATA_FILE, index=False, encoding='utf-8')
            logger.info(f"[METADATA SAVED] Meeting {meeting_id} with attendees: {attendees}")
        except Exception as e:
            logger.error(f"Failed to save meeting metadata: {e}")

    def count_votes_for_meeting(self, meeting_id):
        """Count all votes for each proposed datetime"""
        try:
            if not os.path.exists(RESPONSES_FILE):
                return []

            df = pd.read_csv(RESPONSES_FILE)
            meeting_data = df[df['Meeting ID'] == meeting_id]

            if len(meeting_data) == 0:
                return []

            vote_count = defaultdict(list)

            for _, row in meeting_data.iterrows():
                if row['Status'] == 'responded':
                    try:
                        dt = datetime.strptime(
                            f"{row['Proposed Date']} {row['Proposed Time']}", 
                            '%d-%m-%Y %H:%M'
                        )
                        attendee = row['Attendee']
                        if attendee not in vote_count[dt]:
                            vote_count[dt].append(attendee)
                    except:
                        pass

            sorted_votes = sorted(
                vote_count.items(), 
                key=lambda x: len(x[1]), 
                reverse=True
            )

            return sorted_votes

        except Exception as e:
            logger.error(f"Failed to count votes: {e}")
            return []

    def find_best_meeting_time(self, meeting_id, total_attendees):
        """Find the datetime with most votes (requires 50%+ agreement)"""
        votes = self.count_votes_for_meeting(meeting_id)

        if not votes:
            return None

        best_time = votes[0][0]
        best_votes = votes[0][1]
        agreement_percentage = (len(best_votes) / total_attendees) * 100

        if agreement_percentage >= 50:
            return {
                'datetime': best_time,
                'votes': len(best_votes),
                'agreement': agreement_percentage,
                'attendees': best_votes
            }

        return None

    def send_confirmation_email_english(self, recipient_email, meeting_id, topic, 
                                       best_datetime, all_attendees, organizer,
                                       teams_url=None, google_cal_url=None, outlook_url=None):
        """✅ ENHANCED: Send confirmation with Teams URL + Calendar links"""
        try:
            mail = self.outlook.CreateItem(0)
            mail.To = recipient_email
            mail.Subject = f"✅ {topic} - Meeting Confirmed - {best_datetime.strftime('%d-%m-%Y at %H:%M')}"
            # DO NOT set Categories - causes Outlook error
            # Only use supported properties
            
            day_name = best_datetime.strftime('%A')
            date_str = best_datetime.strftime('%B %d, %Y')
            time_str = best_datetime.strftime('%H:%M')
            
            attendees_str = '\n'.join([f'  • {att}' for att in all_attendees])
            
            mail.Body = f"""Dear {recipient_email.split('@')[0].title()},

✅ MEETING CONFIRMED

Your meeting has been scheduled based on consensus availability.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
MEETING DETAILS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Topic:         {topic}
Date:          {day_name}, {date_str}
Time:          {time_str} (60 minutes)
Organizer:     {organizer}
Meeting ID:    {meeting_id}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ATTENDEES ({len(all_attendees)}):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{attendees_str}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🔗 QUICK LINKS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

            if teams_url:
                mail.Body += f"\n📱 JOIN TEAMS MEETING:\n{teams_url}\n"
            
            if google_cal_url:
                mail.Body += f"\n📅 ADD TO GOOGLE CALENDAR:\n{google_cal_url}\n"
            
            if outlook_url:
                mail.Body += f"\n📅 ADD TO OUTLOOK CALENDAR:\n{outlook_url}\n"
            
            mail.Body += f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
IMPORTANT NOTES:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

✓ This time was automatically selected based on all attendees' availability
✓ Click the Teams link above to join directly (no password needed)
✓ You can also add this meeting to your calendar using the links above
✓ If using Outlook: A calendar invitation is being sent to your inbox
✓ If using Gmail/external email: Click the calendar link to add to your calendar

Please confirm your attendance by replying to this email or joining the Teams meeting.

Thank you!

════════════════════════════════════════════════════════════════════════════════
Automated Meeting Scheduler
Powered by KFSHRC AI Bot
════════════════════════════════════════════════════════════════════════════════
"""
            # AUTO-CLICK: Use send_with_auto_classify to auto-dismiss FORTRA popup
            send_with_auto_classify(mail.Send)
            logger.info(f"[AUTO-CLICK] Confirmation with Teams + Calendar links sent to {recipient_email} (FORTRA popup auto-closed)")
            self.log_timeline(meeting_id, "confirmation_sent", f"English confirmation with Teams + Calendar links sent to {recipient_email}")
            
        except Exception as e:
            logger.error(f"Failed to send confirmation to {recipient_email}: {e}")

    def create_meeting(self, subject, body, start_datetime, duration_minutes, attendee_emails):
        """Create Outlook meeting - only add KFSHRC users to invite"""
        try:
            appt = self.outlook.CreateItem(OL_APPOINTMENT_ITEM)
            appt.MeetingStatus = OL_MEETING
            appt.Subject = subject
            appt.Start = start_datetime
            appt.Duration = duration_minutes
            appt.Body = body

            # Separate KFSHRC from external emails
            kfshrc_emails = [e for e in attendee_emails if '@kfshrc.edu.sa' in e.lower()]
            external_emails = [e for e in attendee_emails if '@kfshrc.edu.sa' not in e.lower()]

            # Add only KFSHRC users to Outlook meeting
            for email in kfshrc_emails:
                try:
                    recipient = appt.Recipients.Add(email)
                    recipient.Type = OL_REQUIRED
                    logger.info(f"Added to Outlook meeting: {email}")
                except Exception as e:
                    logger.warning(f"Could not add to Outlook: {email}: {e}")

            # Try to resolve Outlook recipients
            try:
                appt.Recipients.ResolveAll()
                logger.info("Outlook recipients resolved")
            except:
                logger.info("Some Outlook recipients could not be resolved")
            
            # Send Outlook meeting
            appt.Send()
            logger.info(f"Outlook meeting created: {subject}")
            
            # Log external email recipients
            if external_emails:
                logger.info(f"External email recipients (email only): {', '.join(external_emails)}")
            
            return appt
            
        except Exception as e:
            logger.error(f"Failed to create Outlook meeting: {e}")
            raise

    def save_meeting_log(self, meeting_id, organizer, attendees, topic, duration, scheduled_start, status):
        try:
            log_entry = {
                'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'Meeting ID': meeting_id,
                'Organizer': organizer,
                'Attendees': ', '.join(attendees),
                'Topic': topic,
                'Duration (min)': duration,
                'Scheduled Start': scheduled_start.strftime('%d-%m-%Y %H:%M') if scheduled_start else '',
                'Status': status,
            }

            if os.path.exists(LOG_FILE):
                df = pd.read_csv(LOG_FILE)
                df = pd.concat([df, pd.DataFrame([log_entry])], ignore_index=True)
            else:
                df = pd.DataFrame([log_entry])

            df.to_csv(LOG_FILE, index=False, encoding='utf-8')
            logger.info(f"Meeting logged")
        except Exception as e:
            logger.error(f"Logging error: {e}")

    def process_meeting_requests(self):
        """Process new meeting requests"""
        messages = self.get_unread_messages()
        if len(messages) == 0:
            return

        logger.info(f"Found {len(messages)} message(s)")

        for msg in messages:
            try:
                subject = msg.Subject or ""
                
                if "ID:" not in subject:
                    organizer_email = self.get_sender_smtp_address(msg)
                    body = msg.Body

                    logger.info(f"Processing meeting request: {subject}")
                    logger.info(f"From: {organizer_email}")

                    request = self.extract_meeting_request(body, subject, organizer_email)
                    attendees = request["attendees"]
                    duration = request["duration"]
                    topic = request["topic"]
                    deadline = request["deadline"]

                    if not attendees:
                        reply = msg.Reply()
                        reply.Subject = f"Re: {subject}"
                        reply.Body = """Error: No attendees found.

Please include a line like:
Attendees: john@kfshrc.edu.sa, sarah@gmail.com, external@company.com
Topic: Meeting Topic
Duration: 60
Deadline: 30-07-2026 17:00 (optional)

Note: Accepts any email (KFSHRC, Gmail, external, etc.)
"""
                        reply.Send()
                        msg.UnRead = False
                        msg.Save()
                        continue

                    meeting_id = str(uuid.uuid4())[:8]

                    logger.info(f"Meeting ID: {meeting_id}")
                    logger.info(f"Attendees: {attendees}")
                    logger.info(f"Topic: {topic}")
                    logger.info(f"Deadline: {deadline}")

                    # Mark as read/handled BEFORE the slow, FORTRA-gated sends below
                    # (each mail.Send() blocks on a manual classification click).
                    # Otherwise this message stays "unread" for however long that
                    # takes, and a second bot instance polling the same inbox can
                    # grab it and process it again under a different meeting ID.
                    msg_id = msg.EntryID
                    self.processed_requests.add(msg_id)
                    msg.UnRead = False
                    msg.Save()

                    self.log_timeline(meeting_id, "initiated", f"Meeting request received from {organizer_email}")
                    
                    # SAVE ORIGINAL ATTENDEES so we send confirmation to ALL, not just who responded
                    self.save_meeting_metadata(meeting_id, topic, organizer_email, attendees, deadline)

                    for attendee in attendees:
                        self.send_survey_email_arabic(attendee, meeting_id, topic, organizer_email, deadline)

                    reply = msg.Reply()
                    reply.Body = f"""Meeting request received.

Topic: {topic}
Meeting ID: {meeting_id}
Attendees: {len(attendees)}
Deadline: {deadline.strftime('%d-%m-%Y %H:%M')}

Survey sent to all attendees.
Meeting will be scheduled automatically when consensus is reached.
"""
                    reply.Send()
                    logger.info("Request processed")

            except Exception as e:
                logger.error(f"Error processing request: {e}")

    def process_responses(self):
        """Process attendee responses"""
        messages = self.get_unread_messages()
        if len(messages) == 0:
            return

        for msg in messages:
            try:
                subject = msg.Subject or ""
                
                if "ID:" in subject:
                    body = msg.Body
                    sender = self.get_sender_smtp_address(msg)

                    match = re.search(r'ID:\s*(\S+)\)', subject)
                    if not match:
                        continue

                    meeting_id = match.group(1)

                    logger.info(f"Response received from {sender} for meeting {meeting_id}")

                    proposed_datetimes = self.extract_all_datetimes_from_body(body)

                    if proposed_datetimes:
                        self.save_all_responses(meeting_id, "", sender, proposed_datetimes)
                        self.log_timeline(meeting_id, "response_received", 
                                        f"Response from {sender}: {len(proposed_datetimes)} time options")
                        logger.info(f"Saved {len(proposed_datetimes)} time options from {sender}")
                    else:
                        logger.warning(f"Invalid datetime format in response from {sender}")

                    msg_id = msg.EntryID
                    self.processed_responses.add(msg_id)
                    msg.UnRead = False
                    msg.Save()

            except Exception as e:
                logger.error(f"Error processing response: {e}")

    def check_and_schedule_meetings(self):
        """Check responses and schedule meetings"""
        try:
            if not os.path.exists(RESPONSES_FILE):
                return

            responses_df = pd.read_csv(RESPONSES_FILE)
            
            if len(responses_df) == 0:
                return

            meeting_ids = responses_df['Meeting ID'].unique()

            for meeting_id in meeting_ids:
                try:
                    # Check if already scheduled
                    if os.path.exists(LOG_FILE):
                        log_df = pd.read_csv(LOG_FILE)
                        already_scheduled = log_df[
                            (log_df['Meeting ID'] == meeting_id) & 
                            (log_df['Status'] == 'scheduled')
                        ]
                        
                        if len(already_scheduled) > 0:
                            continue
                    
                    # Get responses for this meeting
                    meeting_responses = responses_df[responses_df['Meeting ID'] == meeting_id]
                    
                    if len(meeting_responses) == 0:
                        continue
                    
                    # Get ORIGINAL attendees from metadata (send confirmation to ALL, not just who responded!)
                    original_attendees = []
                    if os.path.exists(METADATA_FILE):
                        metadata_df = pd.read_csv(METADATA_FILE)
                        meeting_metadata = metadata_df[metadata_df['Meeting ID'] == meeting_id]
                        if len(meeting_metadata) > 0:
                            attendees_str = meeting_metadata.iloc[0]['Original Attendees']
                            original_attendees = [a.strip() for a in attendees_str.split(',')]
                    
                    # Use original attendees, or fall back to who responded
                    if original_attendees:
                        attendees = original_attendees
                        logger.info(f"[ORIGINAL ATTENDEES RETRIEVED] Meeting {meeting_id}: {attendees}")
                    else:
                        # Fallback: use only who responded (old behavior)
                        attendees = list(set(meeting_responses['Attendee'].unique()))
                        logger.warning(f"[METADATA NOT FOUND] Using only respondents: {attendees}")
                    
                    total_attendees = len(attendees)
                    
                    if total_attendees == 0:
                        continue
                    
                    # Check if deadline passed (15 minutes from first response)
                    deadline_passed = False
                    current_time = datetime.now()
                    
                    try:
                        first_response_time = meeting_responses.iloc[0]['Response Time']
                        first_response_dt = datetime.strptime(first_response_time, '%Y-%m-%d %H:%M:%S')
                        time_elapsed = (current_time - first_response_dt).total_seconds() / 60
                        
                        if time_elapsed > 15:
                            deadline_passed = True
                            logger.info(f"Deadline passed for meeting {meeting_id} ({time_elapsed:.0f} minutes)")
                    except:
                        pass
                    
                    # Find best time
                    best_time = self.find_best_meeting_time(meeting_id, total_attendees)
                    
                    # Schedule if: consensus found OR deadline passed
                    if best_time is None and not deadline_passed:
                        continue
                    
                    # If deadline passed but no consensus, pick most popular time
                    if best_time is None and deadline_passed:
                        votes = self.count_votes_for_meeting(meeting_id)
                        if len(votes) > 0:
                            best_time = {
                                'datetime': votes[0][0],
                                'votes': len(votes[0][1]),
                                'agreement': (len(votes[0][1]) / total_attendees) * 100
                            }
                        else:
                            continue
                    
                    logger.info(f"Scheduling meeting {meeting_id}: {best_time['datetime']} with {best_time['votes']} votes")
                    
                    # Get meeting details from METADATA
                    topic = "Meeting"  # Default
                    organizer = "Unknown"  # Default
                    
                    # Try to get topic and organizer from METADATA
                    if os.path.exists(METADATA_FILE):
                        metadata_df = pd.read_csv(METADATA_FILE)
                        meeting_metadata = metadata_df[metadata_df['Meeting ID'] == meeting_id]
                        if len(meeting_metadata) > 0:
                            topic = meeting_metadata.iloc[0]['Topic']
                            organizer = meeting_metadata.iloc[0]['Organizer']
                            logger.info(f"[METADATA] Retrieved topic: {topic}, organizer: {organizer}")
                    
                    # Fallback: Try to get from TIMELINE if metadata not found
                    if topic == "Meeting" and os.path.exists(TIMELINE_FILE):
                        timeline_df = pd.read_csv(TIMELINE_FILE)
                        meeting_timeline = timeline_df[timeline_df['Meeting ID'] == meeting_id]
                        
                        if len(meeting_timeline) > 0:
                            initiated = meeting_timeline[meeting_timeline['Stage'] == 'initiated']
                            if len(initiated) > 0:
                                desc = initiated.iloc[0]['Description']
                                if "from " in desc:
                                    organizer = desc.split("from ")[-1]
                    
                    # Create meeting and send confirmations
                    try:
                        meeting_body = f"Meeting scheduled automatically.\n\nTopic: {topic}"
                        
                        # ENHANCED: Generate Teams + Calendar URLs
                        teams_url = generate_teams_meeting_url(meeting_id, topic)
                        google_cal_url = generate_google_calendar_url(topic, best_time['datetime'], 60, attendees)
                        outlook_url = generate_outlook_calendar_url(topic, best_time['datetime'], 60, attendees)
                        logger.info(f"[URLS GENERATED] Teams + Calendar URLs generated for meeting {meeting_id}")
                        
                        # First: Send confirmations to ALL (this always works)
                        for attendee in attendees:
                            try:
                                self.send_confirmation_email_english(
                                    attendee, 
                                    meeting_id, 
                                    topic, 
                                    best_time['datetime'], 
                                    attendees, 
                                    organizer,
                                    teams_url=teams_url,
                                    google_cal_url=google_cal_url,
                                    outlook_url=outlook_url
                                )
                                logger.info(f"[CONFIRMATION SENT] Email with Teams + Calendar links sent to {attendee}")
                            except Exception as e:
                                logger.error(f"Failed to send confirmation to {attendee}: {e}")
                        
                        # Second: Try to create Outlook meeting (only for KFSHRC users)
                        try:
                            self.create_meeting(topic, meeting_body, best_time['datetime'], 60, attendees)
                            logger.info(f"Outlook meeting created for {meeting_id}")
                        except Exception as e:
                            logger.warning(f"Could not create Outlook meeting for {meeting_id}: {e}")
                            logger.info("But email confirmations were sent to all attendees")
                        
                        # Log meeting
                        self.save_meeting_log(
                            meeting_id, 
                            organizer, 
                            attendees, 
                            topic, 
                            60, 
                            best_time['datetime'], 
                            "scheduled"
                        )
                        
                        # Update timeline
                        self.log_timeline(
                            meeting_id, 
                            "scheduled", 
                            f"Meeting scheduled for {best_time['datetime'].strftime('%d-%m-%Y %H:%M')}"
                        )
                        
                        logger.info(f"Meeting {meeting_id} scheduled and confirmations sent")
                        
                    except Exception as e:
                        logger.error(f"Failed to process meeting {meeting_id}: {e}")
                    
                except Exception as e:
                    logger.error(f"Error processing meeting {meeting_id}: {e}")
                    continue

        except Exception as e:
            logger.error(f"Error in check_and_schedule_meetings: {e}")

    def run(self):
        logger.info("="*80)
        logger.info("BOT STARTED - Intelligent Meeting Scheduler")
        logger.info(f"Checking every {CHECK_INTERVAL} seconds")
        logger.info("="*80)

        try:
            cycle = 0
            while True:
                cycle += 1
                logger.info(f"Cycle {cycle} - {datetime.now().strftime('%H:%M:%S')}")
                self.process_meeting_requests()
                self.process_responses()
                self.check_and_schedule_meetings()
                time.sleep(CHECK_INTERVAL)
        except KeyboardInterrupt:
            logger.info("="*80)
            logger.info("BOT STOPPED")
            logger.info("="*80)


if __name__ == "__main__":
    bot = MeetingBot()
    bot.run()
