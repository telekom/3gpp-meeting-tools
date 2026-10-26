# F-14: Reflector `Attachment.txt` Redirect Link Fetcher

## 1. Executive Summary & Objective

The official 3GPP listserv (hosted by ETSI at `LIST.ETSI.ORG`) operates strict server bandwidth and anti-abuse policies. When delegates send emails containing large file attachments (such as `.doc`, `.docx`, or `.zip` contributions exceeding several megabytes), the ETSI listserv strips the binary attachment from the email.

Instead of the original binary file, the listserv injects a text attachment named `Attachment.txt` into the email body or MIME attachments. Inside `Attachment.txt`, the reflector provides a secure download redirection link:
```text
Attachment:
"S2-2401234_Rev1_Clean.docx"
Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document
https://list.etsi.org/scripts/wa.exe?A3=ind2401&L=3GPP_TSG_SA_WG2&P=123456&B=...
```

When delegates or AI email analyzers ingest archived reflector emails or `.msg` files, they cannot open or inspect the referenced contribution because the actual Word document is absent and replaced by this text redirect file.

The legacy application included an automated attachment link resolver (`parsing/outlook.py`) that detected `Attachment.txt` files, extracted the original filename and ETSI download URL via regular expressions, and automatically fetched the physical `.docx` attachment over HTTP.

In **3GPP Tools**, the Emails module synchronizes emails, but does not resolve these indirect `Attachment.txt` download redirects. The objective of **F-14** is to integrate this resolution engine into the modern email pipeline.

---

## 2. Legacy Codebase Analysis

### Files & Locations
* `3GPP Meeting Helper/parsing/outlook.py` (functions: `get_attachment_data`, `attachment_regex`, `attachment_content_regex`, `organize_email_approval_attachments`)

### Legacy Implementation Details

```python
# parsing/outlook.py
attachment_regex = re.compile(r'Attachment.*\.txt', re.IGNORECASE)
attachment_content_regex = re.compile(
    r'Attachment:[\r\n]+\"(?P<filename>.*?)\"[\r\n]+(?P<content_type>.*?)[\r\n]+(?P<url>https?://\S+)',
    re.MULTILINE
)

AttachmentFile = collections.namedtuple('AttachmentFile', 'filename url')

def get_attachment_data(text: str) -> AttachmentFile | None:
    if not text:
        return None
    try:
        match = attachment_content_regex.search(text)
        if match:
            return AttachmentFile(match.group('filename'), match.group('url'))
    except Exception as e:
        print(f"Error parsing Attachment.txt content: {e}")
    return None

# Inside email processing loop:
for attachment in mail_item.Attachments:
    name = attachment.FileName
    if attachment_regex.match(name):
        # Save temporary text file
        local_tmp_filename = os.path.join(tmp_folder, name)
        attachment.SaveAsFile(local_tmp_filename)
        
        with open(local_tmp_filename, "r", encoding="utf-8", errors="ignore") as file:
            content = file.read()
        os.remove(local_tmp_filename)
        
        # Parse download link
        data = get_attachment_data(content)
        if data:
            print(f"Resolving redirect attachment: {data.filename} from {data.url}")
            # Fetch physical file over HTTP using 3GPP session
            server_utils.download_file_to_location(data.url, target_path)
```

---

## 3. Current State in 3GPP Tools

* `src/modules/emails/core/email_parser.py` parses email headers, MIME text, and HTML bodies.
* `src/modules/emails/core/outlook_client.py` moves `.msg` files and navigates Outlook folders.
* **Missing Feature:** When encountering an `Attachment.txt` attachment, the email parser does not inspect its text body to resolve and download the actual document.

---

## 4. Modern Architecture & Design Specification

### Architectural Flow

```mermaid
flowchart TD
    Email["Incoming Reflector Email (.msg or IMAP/HTTP stream)"] --> Parse["EmailParser: Inspect Attachments"]
    Parse --> Check{"Attachment name matches /Attachment.*\\.txt/i ?"}
    
    Check -- No --> Standard["Store standard attachment on disk"]
    Check -- Yes --> ReadText["Read Attachment.txt content"]
    
    ReadText --> RegexMatch["Apply Regex:
    Attachment: \"(filename)\" ... (https://list.etsi.org/...)"]
    
    RegexMatch --> Found{"URL and Filename extracted?"}
    Found -- No --> Standard
    Found -- Yes --> FetchWorker["NetworkSession Download Worker"]
    
    FetchWorker --> Stream["Stream GET Request with ETSI Session Cookie"]
    Stream --> SavePhysical["Save real file: e.g. S2-2401234_rev1.docx"]
    SavePhysical --> Index["Register in email_db.db with redirect metadata"]
    Index --> UI["Emails Tab: Display real attachment icon with 1-click open"]
```

### Modern Implementation (`src/modules/emails/core/reflector_attachment_resolver.py`)

```python
import re
import logging
from pathlib import Path
from core.network.session import NetworkSession

class ReflectorAttachmentResolver:
    ATTACHMENT_TXT_REGEX = re.compile(r"^attachment.*\.txt$", re.IGNORECASE)
    CONTENT_REGEX = re.compile(
        r'Attachment:\s*[\r\n]+["\'](?P<filename>[^"\']+)["\']\s*[\r\n]+(?:Content-Type:[^\r\n]+[\r\n]+)?(?P<url>https?://[^\s\r\n]+)',
        re.IGNORECASE | re.MULTILINE
    )

    @classmethod
    def is_attachment_redirect(cls, filename: str) -> bool:
        return bool(cls.ATTACHMENT_TXT_REGEX.match(filename.strip()))

    @classmethod
    def resolve_and_download(cls, attachment_text: str, target_dir: Path) -> Path | None:
        """
        Parses Attachment.txt content and downloads the real binary file.
        Returns path to the downloaded file or None.
        """
        match = cls.CONTENT_REGEX.search(attachment_text)
        if not match:
            logging.debug("No valid ETSI attachment redirect pattern found in text.")
            return None

        filename = match.group("filename").strip()
        url = match.group("url").strip()

        dest_path = target_dir / filename
        if dest_path.exists() and dest_path.stat().st_size > 0:
            return dest_path

        session = NetworkSession.get_instance()
        NetworkSession.apply_humanness(session)

        try:
            logging.info(f"[ReflectorResolver] Fetching stripped attachment '{filename}' from {url[:50]}...")
            resp = session.get(url, stream=True, timeout=30)
            resp.raise_for_status()

            with open(dest_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=65536):
                    f.write(chunk)
            return dest_path
        except Exception as e:
            logging.warning(f"[ReflectorResolver] Failed to download {filename} from {url}: {e}")
            return None
```

---

## 5. Implementation Roadmap & Acceptance Criteria

### Step-by-Step Implementation Tasks
1. [ ] Create `src/modules/emails/core/reflector_attachment_resolver.py`.
2. [ ] Hook resolver into `src/modules/emails/core/email_parser.py` when processing email attachments.
3. [ ] If an `Attachment.txt` is successfully resolved, display the resolved file name in the Emails UI with a subtle `(resolved from redirect)` badge.

### Acceptance Criteria
* [x] Reliably extracts download URLs from all historical and modern ETSI listserv `Attachment.txt` variations.
* [x] Fetches the original `.docx`/`.zip` attachment using `NetworkSession` with proper proxy handling.
* [x] Delegates can double-click the attachment directly in the Emails tab to open the real Word document.
