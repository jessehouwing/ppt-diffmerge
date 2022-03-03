# PowerPoint Diff/Merge

A little commandline tool to start PowerPoint in merge mode.

```
ppt-diffmerge-tool "$LOCAL" "$REMOTE" "$BASE" "$RESULT"
```

## Git config

To register this tool in git, add these sections to your git config:

```
[difftool "pptdiffmerge"]
	name = PowerPoint Diff tool
	cmd = C:/ppt-diffmerge/ppt-diffmerge-tool/bin/Debug/ppt-diffmerge-tool.exe "$LOCAL" "$REMOTE"
	binary = true
[mergetool "pptdiffmerge"]
	name = PowerPoint Merge tool
	trustExitCode = false
	keepBackup = false
	cmd = C:/ppt-diffmerge/ppt-diffmerge-tool/bin/Debug/ppt-diffmerge-tool.exe "$LOCAL" "$REMOTE" "$BASE" "$RESULT"
```

Then register the file extensions for PowerPoint by setting the merge tool attributes in your `.gitattributes`:

```
*.ppt	binary diff=pptdiffmerge merge=pptdiffmerge
*.pptm	binary diff=pptdiffmerge merge=pptdiffmerge
*.pptx	binary diff=pptdiffmerge merge=pptdiffmerge
```
