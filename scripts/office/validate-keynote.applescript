on run argv
  set corpusFolder to POSIX file (item 1 of argv) as alias
  set outputFolder to POSIX file (item 2 of argv) as alias
  tell application "Finder" to set sourceFiles to every file of entire contents of corpusFolder whose name extension is "pptx"
  tell application "Keynote"
    repeat with sourceFile in sourceFiles
      set sourcePath to sourceFile as alias
      set presentationDocument to open sourcePath
      set outputPath to ((outputFolder as text) & (name of sourceFile) & ".pdf")
      export presentationDocument to file outputPath as PDF
      close presentationDocument saving no
    end repeat
  end tell
end run

