# `src/outlook_service.py` — Talking to Outlook

## What it does
Wraps the Outlook desktop app (via pywin32 COM). It connects to Outlook, finds
the master template draft, and for each `Contact` creates a **new personalized
draft** — setting the recipient, Cc, subject and the generated table — then
saves it. It **never sends**.

## Why it exists
All the Windows/COM/Outlook-specific complexity lives here, behind a small,
clean interface (`connect`, `locate_template`, `create_draft_for`). The rest of
the app doesn't need to know anything about COM. If you later switch to
Microsoft Graph or Gmail, you'd add a new service and leave everything else
alone.

## Why it is built this way

### pywin32 COM automation
```python
self._app = win32com.client.Dispatch("Outlook.Application")
self._namespace = self._app.GetNamespace("MAPI")
```
This drives the **real, signed-in Outlook** on the PC, so every draft inherits
your actual account, signature and formatting. COM only works **locally** — the
script must run on the same machine as Outlook (this is why automation uses RDP
+ Task Scheduler, not remote calls).

### Lazy import of `win32com`
```python
def connect(self):
    import win32com.client
```
The import is inside `connect()`, not at the top of the file. That lets the
module be imported and unit-tested on **any** OS (e.g. for the table/insert
logic) without pywin32 installed. The Outlook dependency is only needed when you
actually connect.

### Finding the template: subject or EntryID
```python
if self.template_entryid: ... GetItemFromID(...)
else: scan Drafts for matching Subject
```
- By **subject** (`MASTER TEMPLATE`) is easy and human-friendly.
- By **EntryID** is exact and unambiguous if you ever have two drafts with the
  same subject. Configurable via `TEMPLATE_ENTRYID`.

### Duplicating via an `.oft` template — the key robustness fix
```python
self._template.SaveAs(path, OL_SAVE_AS_TEMPLATE)   # cache once
...
self._app.CreateItemFromTemplate(self._oft_path)   # fresh copy per row
```
**Why not just `template.Copy()`?** If your master draft was created with
*Reply* or *Forward*, Outlook marks it an "inline response item" and `Copy()`
fails with *"This method can't be used with an inline response mail item."*
Saving the template once as a `.oft` and creating each draft from it sidesteps
that limitation **and** preserves fonts, images, hyperlinks and the signature.
If saving the `.oft` fails for any reason, it falls back to `Copy()`
automatically.

### Only three things change per row
```python
new_mail.To = contact.recipient                 # column A
if self.cc_address: new_mail.CC = self.cc_address
new_mail.Subject = contact.subject(self.subject_columns)  # first 7 columns
new_mail.HTMLBody = self._insert_table(new_mail.HTMLBody, table_html)
```
Everything else (layout, signature, embedded images) comes untouched from the
template copy — meeting the "preserve existing formatting" requirement.

### Table insertion: placeholder or append
```python
if self.table_placeholder in body:
    return body.replace(self.table_placeholder, table_html)
idx = body.lower().rfind("</body>")
return body[:idx] + table_html + body[idx:]   # else insert before </body>
```
- If the draft contains `{{TABLE}}`, the table goes exactly there.
- Otherwise it's inserted just before `</body>` so it appears at the end while
  keeping the signature and existing HTML valid.

### `Save()`, never `Send()`
```python
new_mail.Save()  # saved as a draft; never .Send()
```
The single most important safety guarantee in the project. There is no `.Send()`
call anywhere.

### Returns the new draft's EntryID
Useful for logging and for future features like duplicate detection or building
a report of what was created.

## Outlook constants used
| Constant | Value | Meaning |
|----------|-------|---------|
| `OL_FOLDER_DRAFTS` | 16 | The Drafts default folder |
| `OL_SAVE_AS_TEMPLATE` | 2 | Save a mail item as a `.oft` template |
