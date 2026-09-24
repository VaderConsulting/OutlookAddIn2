# OutlookAddIn2

MailInEvent is a Visual Studio 2008 VSTO Outlook 2007 add-in that handles `Application.NewMail`. On a new message it reads the Inbox item, then shows a message box with the sender address and a subject line that includes the sender name. It is my working copy from Historical Dev, published so the mail-event experiment can be opened from GitHub.

**Source last updated:** 2008-02-26  
**Language:** VB.NET  
**Target:** .NET 3.5 / VSTO Outlook 2007  
**Output:** Outlook add-in library (`MailInEvent`)

## Solution structure

| Project | Language | Type | Purpose |
|---------|----------|------|---------|
| `MailInEvent` | VB.NET | VSTO library (`OutlookAddIn2/MailInEvent.vbproj`) | `ThisAddIn` NewMail handler; MsgBox of sender and subject |

## How to open

Open `OutlookHandler.sln` in Visual Studio 2008 with VSTO and Outlook 2007. The ClickOnce manifest is signed with `OutlookAddIn2_TemporaryKey.pfx` (gitignored). Generate a new test cert if you need to build.

## Requirements

- Visual Studio 2008, .NET Framework 3.5

## Attribution and provenance

From Dave Robinson's Historical Dev archive (OneDrive folder `OutlookAddIn2`). Assembly title OutlookAddIn2, copyright 2008. No third-party libraries in the tree.

## License

MIT License. Copyright (c) 2026 VaderConsulting.
