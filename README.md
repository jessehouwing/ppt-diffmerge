# PowerPoint Diff/Merge

A little commandline tool to start PowerPoint in merge mode.

```
ppt-diffmerge-tool --local="$LOCAL" --remote="$REMOTE" --base="$BASE" --output="$RESULT" 
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

Then register the file extensions for PowerPoint by setting the merge tool attroibutes in your `.gitattributes`:

```
*.ppt	binary diff=ppt-diffmerge merge=ppt-diffmerge
*.pptm	binary diff=ppt-diffmerge merge=ppt-diffmerge
*.pptx	binary diff=ppt-diffmerge merge=ppt-diffmerge
```
