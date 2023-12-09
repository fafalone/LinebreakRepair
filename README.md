# LinebreakRepair
GHFix: VB6 GitHub Linebreak Repair

![image](https://github.com/fafalone/LinebreakRepair/assets/7834493/48230542-0803-46f6-8004-5c490d226015)

I've been downloading a lot of VB6 projects from GitHub to test in twinBASIC. When these have been uploaded with certain settings, GitHub replaces the vbCrLf line breaks with vbLf only, which results in a corrupt file that VB6 can't read. tB can-- but I need to confirm working in VB6 first. So I made this small project to automatically repair the line breaks of all VB6 files in a given directory, with a number of options to help. The file type list is preset, but you could change it to work on any file type instead.

This project uses my tbShellLib project for all the APIs (this is why the source file size is so large; however the compiled exe uses only what is neccessary, so is only 2MB). Can be compiled to both 32bit and 64bit. I've tested to confirm the output is byte for byte identical to using WinHex to manually replace 0x0A with 0x0D 0x0A, and the app checks whether the line breaks already appear to be correct.


### How it works

The basic principle is two loops. First, we count how many Lf's (0x0A) there are while checking if there's already an 0x0D before it:

```vb6

            For j = 0 To UBound(btSrc)
                If btSrc(j) = btFind Then
                    If (nMatch = 0) And (bFlagOk = False) And (j <> 0) Then
                        If btSrc(j - 1) = btApp Then
                            Dim r As VbMsgBoxResult = MsgBox("Warning: The current file appears to already have correct line breaks; continue anyway?" & vbCrLf & "Press 'Try again' to skip to the next file, 'Continue' to proceed and suppress future warnings, or 'Cancel' to abort the entire operation.", vbCritical Or vbCancelTryAgainContinue, sName)
                            If r = vbCancel Then
                                CloseHandle hFile
                                Return 1
                            ElseIf r = vbContinue Then
                                bFlagOk = True
                            ElseIf r = vbTryAgain Then
                                CloseHandle hFile
                                GoTo NextFile
                            End If
                        End If
                    End If
                    cbDst += 1
                    nMatch += 1
                End If
            Next
```

This is used to compute the total size we'll need for the repaired file, so we can copy it right in with ease:

```vb6
            ReDim btDst(cbDst - 1)
            
            
            For j = 0 To UBound(btSrc)
                If btSrc(j) = btFind Then
                    btDst(k) = btApp
                    btDst(k + 1) = btFind
                    k += 2
                Else
                    btDst(k) = btSrc(j)
                    k += 1
                End If
            Next
```

Then we just make sure the file pointer is at the beginning and save:

```vb6
            lRet = SetFilePointer(hFile, 0, ByVal 0, FILE_BEGIN)
            If (lRet = INVALID_SET_FILE_POINTER) Or (Err.LastDllError <> NO_ERROR) Then
                PostLog "Failed to reset file pointer on " & sName
                CloseHandle hFile
                Return 2
            End If
            lRet = WriteFile(hFile, btDst(0), cbDst, cbRet, vbNullPtr)
            
            If lRet Then
                PostLog "Successfully fixed " & nMatch & IIf(nMatch = 1, " line break", " line breaks") & " in " & sName
            Else
                PostLog "Failed to write output to " & sName & ", " & PrintError(Err.LastDllError)
            End If
            
            CloseHandle hFile
```

And that's the core of it. The snippets above are just to illustrate technique, there's a lot of missing error checking and other stuff. Download the project or browse Export for the full source.
