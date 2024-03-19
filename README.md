This Outlook macro exports Outlook Notes to Joplin.

Instructions
============

1. Open Outlook for Microsoft 365. May work with other versions, but macros are reportedly not supported in the [New Outlook for Windows](https://support.microsoft.com/en-gb/office/getting-started-with-the-new-outlook-for-windows-656bb8d9-5a60-49b2-a98b-ba7822bc7627).
2. See these instructions to [create a macro with Microsoft Visual Basic for Applications (VBA)](https://support.microsoft.com/en-gb/office/create-a-macro-in-outlook-ffc49e8c-0e5b-4daa-912d-e68c6c46bf27)
3. In VBA `Tools` > `References`, check that you have `Microsoft Scripting Runtime` enabled.
4. In [`Outlook Notes to Joplin.bas`](https://gist.github.com/timadye/c0cd594f08c6b1d6a2c8d48be396da56#file-outlook-notes-to-joplin-bas), in this Gist below, select `Raw` mode and Copy the text.
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