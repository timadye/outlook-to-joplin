# Outlook Notes to Joplin

This Outlook macro exports Outlook Notes to Joplin.

It is based on [@ramisedhom's macro](https://gist.github.com/ramisedhom/0f34c5d6a8d73f0b98ac4bea2ec30be0) for exporting Outlook Mail, which was also independently [developed by @breisfeld](https://gist.github.com/breisfeld/af22feeab3ba0849a9fb6c7ab596992b).

## Instructions

1. Open Outlook for Microsoft 365. May work with other versions, but macros are reportedly not supported in the [New Outlook for Windows](https://support.microsoft.com/en-gb/office/getting-started-with-the-new-outlook-for-windows-656bb8d9-5a60-49b2-a98b-ba7822bc7627).
2. See these instructions to [create a macro with Microsoft Visual Basic for Applications (VBA)](https://support.microsoft.com/en-gb/office/create-a-macro-in-outlook-ffc49e8c-0e5b-4daa-912d-e68c6c46bf27)
3. In VBA `Tools` > `References`, check that you have `Microsoft Scripting Runtime` enabled.
4. Copy the text from [`Outlook Notes to Joplin.bas`](https://gist.github.com/timadye/c0cd594f08c6b1d6a2c8d48be396da56#file-outlook-notes-to-joplin-bas) to the clipboard. You can select `Raw` mode to get the text without formatting.
5. Paste the text into `ThisOutlookSession` window in VBA
6. Download [`JSON.bas`](https://github.com/omegastripes/VBA-JSON-parser/blob/master/JSON.bas) to your computer
7. In VBA: `File` > `Import File...` and select your downloaded `JSON.bas`
8. Open Joplin
9. In Joplin: `Tools` > `Options` > `Web Clipper`. Select `Enable Web Clipper Service`
10. On the same page, under Authorisation token, select `copy Token`
11. Back in VBA, replace the text `REPLACE ME WITH YOUR TOKEN` with the copied token (inside the quotes, `"a1b2..."`).
12. Save you VBA project (Ctrl+S) and close VBA.
13. Open the list of Notes in Outlook, and select one or more notes that you want to export to Joplin. Use Ctrl+A to select all notes.
14. Run the Macro `Project1.ThisOutlookSession.SendToJoplin`. See the [instructions to run a macro in Outlook](https://support.microsoft.com/en-gb/office/run-a-macro-in-outlook-2e03e2e5-e637-4416-9ea0-2230151b0c31).
15. When it is done, a message box will open saying how many notes were exported to Joplin.
16. The notes will be in a Joplin workbook called `Outlook Notes`. Since Outlook notes are plain text, they will be imported as Markdown source. The `Created`, `Modified`, and `Categories` fields will be exported to Joplin `Created`, `Updated`, and `Tags` fields.

## Discussion

* [Export Outlook Notes to Joplin (19/03/2024)](https://discourse.joplinapp.org/t/export-outlook-notes-to-joplin/36901) - discussion of this Gist
* [First trial to create note from Outlook (22/12/2019)](https://discourse.joplinapp.org/t/first-trial-to-create-note-from-outlook/4822)
* [Import Microsoft Outlook notes (25/04/2020)](https://discourse.joplinapp.org/t/import-microsoft-outlook-notes/8201)
* [Export email from Outlook to Joplin (14/04/2022)](https://discourse.joplinapp.org/t/export-email-from-outlook-to-joplin/25148)

## Revisions

19/03/2023 Allow special characters in title. Account for timezones. Cache tag ids. Simpler search.
