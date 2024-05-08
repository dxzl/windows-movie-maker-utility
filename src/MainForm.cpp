//---------------------------------------------------------------------------
// Software by Scott Swift - This program is distributed under the
// terms of the GNU General Public License. Use and distribute freely.
//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop

#include "MainForm.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"

// By Scott M. Swift www.yahcolorize.com,
// dxzl@live.com June 15, 2017
// Version 1.1 August 26, 2018
// Version 1.2 March 11, 2019
// Version 1.4 April 11, 2024

#define MAX_DISPLAY_FILES 25 // max files to list in a ShowMessage()
#define TIMEOUT_SEARCH 10000 // 10 sec search timeout (mS)

TFormMain *FormMain;

// callback function
INT CALLBACK BrowseCallbackProc(HWND hwnd, UINT uMsg, LPARAM lp, LPARAM pData)
{
    if (uMsg==BFFM_INITIALIZED) SendMessage(hwnd, BFFM_SETSELECTION, TRUE, pData);
    return 0;
}

//---------------------------------------------------------------------------
__fastcall TFormMain::TFormMain(TComponent* Owner)
  : TForm(Owner)
{
  GMediaFolderPath = "";
  GProjectFileName = "";
  GCommonPath = "";
  GSearchFileName = "";
  GbQuit = false;
  GbQuitAll = false;

  //enable drag&drop files
  ::DragAcceptFiles(this->Handle, true);
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::FormShow(TObject *Sender)
{
  String s = "FIRST: Make a backup of your .wlmp Windows Movie Maker project file!\n\n"
    "Follow the steps on the buttons, 1,2,3,4.\n\n"
    "(Hint: For steps 1 and 2 you can drag-drop the project-file into this window.\n"
    "Then drag the folder with your movie and image files into this window.\n"
    "Finally, press \"Apply new root-project path\" and press \"Save file\")\n\n"
    "Cheers, Scott Swift (dxzl@live.com)";
  Memo1->Lines->Text = s;
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::TimerSearchTimeoutTimer(TObject *Sender)
{
  TimerSearchTimeout->Enabled = false;
  String sMsg = "It is taking a long time to find \"" + GSearchFileName +
    "\". Click YES to quit this search, NO to continue or CANCEL to abort.";

  int mbResult = MessageBox(Handle, sMsg.w_str(), L"Quit searching?",
            MB_ICONQUESTION + MB_YESNOCANCEL + MB_DEFBUTTON2);

  if (mbResult == IDYES)
    GbQuit = true;
  else if (mbResult == IDCANCEL){
    GbQuit = true;
    GbQuitAll = true;
  }
  else // restart timer
    TimerSearchTimeout->Enabled = false;
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::ButtonApplyNewRootPathClick(TObject *Sender)
{
  if (GProjectFileName.IsEmpty()){
    ShowMessage("Please click \"1. Read .wlmp project-file\"");
    return;
  }

  if (GMediaFolderPath.IsEmpty()){
    ShowMessage("Please click \"2. Select folder where your image and "
       "movie/image/audio files are located\"");
    return;
  }

  TStringList* slMissingFiles = NULL;

  try
  {
    slMissingFiles = new TStringList();

    try
    {
      int iCount = Memo1->Lines->Count;
      if (!iCount)
        return;

      int iFilePathCount = GetFilePathCount();
      if (!iFilePathCount)
        return;

      ProgressBar1->Position = 0;
      ProgressBar1->Step = 1;
      ProgressBar1->Min = 0;
      ProgressBar1->Max = iFilePathCount;

      GCommonPath = GetCommonPath(); // get again (in case edit text manually changed)

      int lenOldCommonPath = GCommonPath.Length();

      GbQuitAll = false;

      for(int ii = 0 ; ii < iCount ; ii++)
      {
        String OldPath;
        String OldStr = Memo1->Lines->Strings[ii];
        int Pos1 = OldStr.Pos("filePath=\"");
        if (Pos1 == 0)
          continue;

        ProgressBar1->StepIt();

        // get the old path and find the trailing quote
        int len = OldStr.Length();
        int jj;
        for (jj = Pos1+10; jj <= len; jj++)
        {
          Char c = OldStr[jj];
          if (c == '\"') break;
          OldPath += c;
        }
        if (jj > len) // no trailing quote, continue
          continue;

        OldPath = XmlDecode(OldPath);

  //      String FileName = ExtractFileName(OldPath);
  //      String FilePath = ExtractFilePath(OldPath);
        String sRemainingPath = OldPath.SubString(lenOldCommonPath+1, OldPath.Length());

        // Need to remove the first subdirectory if it does not exist...
        // It's being replaced with the new folder we drag-dropped...
        // So if there is a '\' or '/', remove text up to and including the first one
        int Pos2 = sRemainingPath.Pos("\\");
        if (Pos2 == 0)
          Pos2 = sRemainingPath.Pos("/");
        if (Pos2 != 0)
          if (!DirectoryExists(ExtractFilePath(GMediaFolderPath + sRemainingPath)))
            sRemainingPath = sRemainingPath.SubString(Pos2+1, sRemainingPath.Length()-Pos2);

        String sFullPath = GMediaFolderPath + sRemainingPath;

        if (!FileExists(sFullPath))
        {
          String sFind = RecurseFind(GMediaFolderPath, ExtractFileName(sFullPath));

          if (GbQuitAll){
            Memo1->Clear();
            ProgressBar1->Position = 0;
            GMediaFolderPath = "";
            GProjectFileName = "";
            GCommonPath = "";
            GSearchFileName = "";
            GbQuit = false;
            GbQuitAll = false;
            LabelMediaFolderPath->Caption = "";
            LabelProjectFilePath->Caption = "";
            ShowMessage("Aborted. Try choosing a media files folder "
                               "that contains fewer sub-folders...");
            return;
          }

          if (sFind.IsEmpty()){
            slMissingFiles->Add(sFullPath);
            sFullPath = "";
          }
          else
            sFullPath = sFind;
        }

        if (!sFullPath.IsEmpty()) {
          // replace oldpath with newpath
          String NewStr = OldStr.SubString(1, Pos1+10-1) + XmlEncode(sFullPath);

          // add the rest of the original line
          for (; jj <= len; jj++)
            NewStr += OldStr[jj];

          Memo1->Lines->Strings[ii] = NewStr;
        }
      }

      // Move to 3rd line, Character 3:
  //    Memo1->SelStart = Memo1->Perform(EM_LINEINDEX, 2, 0) + 3;
  //    Memo1->SelLength = 0;
  //    Memo1->Perform(EM_SCROLLCARET, 0, 0);
  //    Memo1->SetFocus();

      // Move to line 0, Character 0:
      Memo1->SelStart = Memo1->Perform(EM_LINEINDEX, 0, 0);
      Memo1->SelLength = 0;
      Memo1->Perform(EM_SCROLLCARET, 0, 0);
      Memo1->SetFocus();
    }
    catch(...)
    {
      ShowMessage("Threw exception...");
    }
  }
  __finally{
    if (slMissingFiles){
      if (slMissingFiles->Count)
        ShowMessage("Missing media files: \n\n" + slMissingFiles->Text);
      delete slMissingFiles;
    }
    ProgressBar1->Position = 0;
  }
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::CopyOrMoveProject(bool bMove)
{
  // Copy/Move movies/images for project to new folder and generate
  // a new project file to access them.

  Memo1->Clear();

  LabelMediaFolderPath->Caption = "";
  GMediaFolderPath = "";

  LabelProjectFilePath->Caption = "";
  GProjectFileName = "";

  GCommonPath = "";

  String sMoveOrCopy = bMove ? "MOVE" : "COPY";

  // 1. Load the movie-maker project file into Memo1
  ShowMessage("Click OK to select a Windows Live Movie Maker "
       "project file (.wlmp) for " + sMoveOrCopy + "...");
  ButtonReadProjectFileClick(NULL);
  if (Memo1->Lines->Count == 0)
    return;

  // 2. Get destination folder for copied media files

  // GMediaFolderPath must be set to the path of the project file before calling
  // SelectFolder()
  String sProjectFilePath = ExtractFilePath(GProjectFileName);

  String NewMediaFilesPath = BrowseForFolder(Handle,
    "Choose a destination folder to " + sMoveOrCopy + " this "
    "movie's media files into...", sProjectFilePath);

  if (NewMediaFilesPath.IsEmpty())
    return;

  NewMediaFilesPath = IncludeTrailingPathDelimiter(NewMediaFilesPath);

  String sMsg = "Click NO if you have not backed up your original movie-project and media files!\n\n"
    "Click YES to " + sMoveOrCopy + " media files. Choose NO to cancel.";
  String sMsg2 = sMoveOrCopy + " media files?";
  if (MessageBox(Handle, sMsg.w_str(), sMsg2.w_str(),
            MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON2) == IDNO)
    return;

  // 3. process media file names
  TStringList* pSlSourceNoExist = NULL;
  TStringList* pSlCopySuccess = NULL;
  TStringList* pSlFailDelete = NULL;

  ProgressBar1->Position = 0;
  ProgressBar1->Step = 1;
  ProgressBar1->Min = 0;

  try{
    try{
      int iFilepathCount = GetFilePathCount();
      if (!iFilepathCount)
        return;

      pSlSourceNoExist = new TStringList();
      pSlCopySuccess = new TStringList();
      pSlFailDelete = new TStringList();

      if (!pSlSourceNoExist || !pSlCopySuccess || !pSlFailDelete)
        return;

      ProgressBar1->Max = iFilepathCount;

      for(int ii = 0; ii < Memo1->Lines->Count; ii++){
        String OldFullPath;
        String OldStr = Memo1->Lines->Strings[ii];
        int Pos1 = OldStr.Pos("filePath=\"");
        if (Pos1 == 0)
          continue;

        ProgressBar1->StepIt();

        // get the old path and find the trailing quote
        int len = OldStr.Length();
        int jj;
        for (jj = Pos1+10; jj <= len; jj++){
          Char c = OldStr[jj];
          if (c == '\"') break;
          OldFullPath += c;
        }
        if (jj > len) // no trailing quote, continue
          continue;

        OldFullPath = XmlDecode(OldFullPath);

        String FileName = ExtractFileName(OldFullPath);
        String NewFullPath = NewMediaFilesPath + FileName;

        if (OldFullPath == NewFullPath)
          // if the old project-file is already at the destination, copy will fail!
          pSlCopySuccess->Add(OldFullPath);
        else if (FileExists(OldFullPath)){
          try{
            TFile::Copy(OldFullPath, NewFullPath, true); // set overwrite true!
            pSlCopySuccess->Add(OldFullPath);
            Application->ProcessMessages();
          }
          catch(...){
            ShowMessage("Aborting! Unable to copy \"" + OldFullPath +
                        "\" to \"" + NewFullPath + "\"");
            return;
          }
        }
        else // file in project-file does not exist where expected
          pSlSourceNoExist->Add(OldFullPath);

        if (OldFullPath != NewFullPath){
          String sNewLine = OldStr.SubString(1, Pos1+10-1);
          sNewLine += XmlEncode(NewFullPath);
          sNewLine += OldStr.SubString(jj, OldStr.Length());
          Memo1->Lines->Strings[ii] = sNewLine;
        }
      }

      // Move to line 0, Character 0:
      Memo1->SelStart = Memo1->Perform(EM_LINEINDEX, 0, 0);
      Memo1->SelLength = 0;
      Memo1->Perform(EM_SCROLLCARET, 0, 0);
      Memo1->SetFocus();

      if (pSlCopySuccess->Count != iFilepathCount-pSlSourceNoExist->Count){
        ShowMessage("Aborting. Failed to copy " +
          String((iFilepathCount-pSlSourceNoExist->Count)-pSlCopySuccess->Count) +
                  " file(s)!");
        return;
      }

      String s1 = "Successfully copied " + String(pSlCopySuccess->Count) +
                     " media file(s)!\n\n";
      if (bMove){
        s1 += "You chose to MOVE the project - are you sure you want to "
         "delete copied media files from their original locations? Choose NO "
         "unless you have a backup of these files!";

        int button = MessageBox(Handle, s1.w_str(), L"Delete Media Files?",
                  MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON2);

        if (button == IDYES){
          for (int ii=0; ii < pSlCopySuccess->Count; ii++){
            try{
              TFile::Delete(pSlCopySuccess->Strings[ii]);
            }
            catch(...){
              pSlFailDelete->Add(pSlCopySuccess->Strings[ii]);
            }
            Application->ProcessMessages();
          }

          if (pSlFailDelete->Count){
            s1 += "Unable to delete " + String(pSlFailDelete->Count) +
             " files after copying! (could be several reasons...)\n\n";

            for (int ii=0; ii < pSlFailDelete->Count && ii < MAX_DISPLAY_FILES; ii++)
              s1 += pSlFailDelete->Strings[ii] + "\n";

            s1 += "\n";
          }
        }
      }

      if (pSlSourceNoExist->Count){
        s1 += "The project file references " + String(pSlSourceNoExist->Count) +
         " file(s) that don't exist in the specified location:\n\n";

        for (int ii=0; ii < pSlSourceNoExist->Count && ii < MAX_DISPLAY_FILES; ii++)
          s1 += pSlSourceNoExist->Strings[ii] + "\n";

        s1 += "\nYou can choose \"Cancel\" in the next dialog or you can save "
          "the project using a new name and give it a try!";
      }

      if (!s1.IsEmpty())
        ShowMessage(s1);

      // 4. save new project file with paths pointing to copied/moved media files
      // before calling ButtonSaveFileClick(NULL) we expect GProjectFileName
      // to have the full project file path and name!
      ShowMessage("Next, you will choose where to save the newly modified "
           "movie project file. Click OK to continue...");
      DoSaveFileDialog();
    }
    catch(...){
      ShowMessage("Exception thrown!");
    }
  }
  __finally{
    if (pSlSourceNoExist)
      delete pSlSourceNoExist;
    if (pSlCopySuccess)
      delete pSlCopySuccess;
    if (pSlFailDelete)
      delete pSlFailDelete;

    ProgressBar1->Position = 0;
  }

}
//---------------------------------------------------------------------------
int __fastcall TFormMain::GetFilePathCount()
{
  int iFilepathCount = 0;
  for(int ii = 0; ii < Memo1->Lines->Count; ii++){
    String OldFullPath;
    String OldStr = Memo1->Lines->Strings[ii];
    int Pos1 = OldStr.Pos("filePath=\"");
    if (Pos1 != 0)
      iFilepathCount++;
  }
  return iFilepathCount;
}
//---------------------------------------------------------------------------
String __fastcall TFormMain::GetCommonPath()
{
  TStringList *slPaths = NULL;
  String sOldCommonPath;

  try{
    slPaths = new TStringList();

    int iCount = Memo1->Lines->Count;

    // First, get all the old paths into a stringlist
    for(int ii = 0 ; ii < iCount ; ii++){
      String OldPath;
      String OldStr = Memo1->Lines->Strings[ii];
      int Pos1 = OldStr.Pos("filePath=\"");
      if (Pos1 == 0)
        continue;

      // get the old path and find the trailing quote
      int len = OldStr.Length();
      int jj;
      for (jj = Pos1+10; jj <= len; jj++){
        Char c = OldStr[jj];
        if (c == '\"') break;
        OldPath += c;
      }
      if (jj > len) // no trailing quote, continue
        continue;

      slPaths->Add(ExtractFilePath(OldPath));
    }

    sOldCommonPath = CommonPath(slPaths);
  }
  __finally{
    if (slPaths) delete slPaths;
  }
  return sOldCommonPath;
}
//---------------------------------------------------------------------------
String __fastcall TFormMain::CommonPath(TStringList *slPaths)
{
    String sPath;

    int count = slPaths->Count;
    if (count < 2) return ""; // need two strings to compare

    // get "min" path length
    int min = 10000;
    for (int ii=0; ii < count; ii++){
      int i = slPaths->Strings[ii].Length();
      if (i < min)
        min = i;
    }

    // compare up to "min" chars for every path
    for (int jj=1; jj <= min; jj++){
      Char c1 = slPaths->Strings[0][jj];
      for (int ii=1; ii < count; ii++){
        Char c2 = slPaths->Strings[ii][jj];
        if (c1 != c2)
          return sPath;
      }
      sPath += c1;
    }
    return sPath;
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::ButtonHelpClick(TObject *Sender)
{
  String sHelp = "Windows Live Movie Maker project-file path corrector\n"
                                      "by Scott Swift 2024\n\n";
  sHelp += "The Problem: Windows Live Movie Maker project files use absolute file-paths.\n"
      "That means that after you make a movie, you can't change the location "
      "of the media files in your movie. This app solves that problem!\n\n"
      "Use the 1,2,3,4 button steps when you have an old Windows Live Movie Maker project "
      "that no longer works because your media files have moved.\n\n"
      "Additionally, \"Tools->Copy/Move project to new folder\" lets you move a working project, "
      "perhaps using media files located in many places on your computer, and consolidate/copy "
      "all of those media files into a single folder. Then you can save the new project file anywhere.";
  ShowMessage(sHelp);
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::ButtonSelectMediaFolderClick(TObject *Sender)
{
  String sMediaFilesPath = BrowseForFolder(Handle,
         "Choose the top-level folder where your media files for "
                "this project are located...", GMediaFolderPath);

  if (sMediaFilesPath.IsEmpty())
    return;

  GMediaFolderPath = IncludeTrailingPathDelimiter(sMediaFilesPath);
  LabelMediaFolderPath->Caption = "Media-files: \"" + GMediaFolderPath + "\"";
}
//---------------------------------------------------------------------------
String __fastcall TFormMain::SelectFolder(String sCaption, String sPath)
{
    UnicodeString selDir;
    TSelectDirExtOpts opt;

    opt.Clear();
    opt += TSelectDirExtOpts()<<sdNewFolder;
    opt += TSelectDirExtOpts()<<sdShowEdit;
    opt += TSelectDirExtOpts()<<sdNewUI;

    if (SelectDirectory(sCaption.c_str(), L"", sPath, opt, this ))
        return sPath;
    return "";
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::ButtonReadProjectFileClick(TObject *Sender)
{
  OpenDialog1->Title = "Open File";
  OpenDialog1->DefaultExt = "wlmp";
  OpenDialog1->Filter = "Windows Movie Maker files (*.wlmp)|*.wlmp|"
           "All files (*.*)|*.*";
  OpenDialog1->FilterIndex = 1; // start the dialog showing ini files
  OpenDialog1->Options.Clear();
  OpenDialog1->Options << ofHideReadOnly
  << ofPathMustExist << ofFileMustExist << ofEnableSizing;

  if (!OpenDialog1->Execute()){
    Memo1->SetFocus();
    return; // Cancel
  }

  GProjectFileName = OpenDialog1->FileName;
  LabelProjectFilePath->Caption = "Project file: \"" + GProjectFileName + "\"";

  LoadFile();
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::ButtonSaveFileClick(TObject *Sender)
{
  if (GProjectFileName.IsEmpty()){
    ShowMessage("Please click \"1. Read .wlmp project-file\"");
    return;
  }

  if (GMediaFolderPath.IsEmpty()){
    ShowMessage("Please click \"2. Select folder where your image and "
       "movie/image/audio files are located\"");
    return;
  }

  DoSaveFileDialog();
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::DoSaveFileDialog(void)
{
  try{
    SaveDialog1->Title = "Save File";
    SaveDialog1->DefaultExt = "wlmp";
    SaveDialog1->Filter = "Windows Movie Maker files (*.wlmp)|*.wlmp|"
             "All files (*.*)|*.*";
    SaveDialog1->FilterIndex = 2; // start the dialog showing all files
    SaveDialog1->Options.Clear();
    SaveDialog1->Options << ofHideReadOnly
     << ofPathMustExist << ofOverwritePrompt << ofEnableSizing
        << ofNoReadOnlyReturn;
    SaveDialog1->FileName = GProjectFileName;

    if (SaveDialog1->Execute()){
      GProjectFileName = SaveDialog1->FileName;
      LabelProjectFilePath->Caption = "Project file: \"" + GProjectFileName + "\"";
      Memo1->Lines->SaveToFile(GProjectFileName, TEncoding::UTF8);
    }
  }
  catch(...){
    ShowMessage("Can't save file: \"" + SaveDialog1->FileName + "\"");
  }

  Memo1->SetFocus();
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::WMDropFile(TWMDropFiles &Msg)
{
    wchar_t* wBuf = NULL;

    try{
        //get dropped files count
        int cnt = ::DragQueryFileW((HDROP)Msg.Drop, -1, NULL, 0);

        if (cnt == 0 || cnt > 2) return;

        wBuf = new wchar_t[MAX_PATH];
        for (int ii=0; ii < cnt; ii++){
          // Get next file-name
          if (::DragQueryFileW((HDROP)Msg.Drop, 0, wBuf, MAX_PATH) > 0){
              // Load and convert file as per the file-type (either plain or rich text)
              WideString wFile(wBuf);

              // don't process this drag-drop until previous one sets m_DragDropFilePath = ""
              if (!wFile.IsEmpty()){
                  String sFile = String(wFile); // convert to utf-8 internal string

                  if (FileExists(sFile)){ // is it a file? (then must be the project-file!)
                    // handle drag-drop of the .wlmp movie project-file
                    GProjectFileName = sFile;
                    LabelProjectFilePath->Caption = "Project file: \"" + GProjectFileName + "\"";
                    LoadFile();
                  }
                  else{ // is it a directory? (then must be the folder that has our photos and video clips!)
                    if (DirectoryExists(sFile)){
                      // handle drag-drop of the folder with video clips and photos
                      GMediaFolderPath = IncludeTrailingPathDelimiter(sFile);
                      LabelMediaFolderPath->Caption = "Media-files: \"" + GMediaFolderPath + "\"";
                    }
                  }
              }
          }
        }
    }
    __finally{
      try { if (wBuf != NULL) delete [] wBuf; } catch(...) {}
    }
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::LoadFile(void)
{
  try{
    Memo1->Clear();

    // Load wlmp file
    Memo1->Lines->LoadFromFile(GProjectFileName, TEncoding::UTF8);

    GCommonPath = GetCommonPath();
  }
  catch(...){
    ShowMessage("Can't load file: \"" + GProjectFileName + "\"");
  }

  Memo1->SetFocus();
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::Moveprojecttonewfolder1Click(TObject *Sender)
{
  CopyOrMoveProject(true);
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::CopyMovieImageFilesToNewFolder1Click(TObject *Sender)
{
  CopyOrMoveProject(false);
}
//---------------------------------------------------------------------------
String __fastcall TFormMain::XmlDecode(String sIn)
{
  // XML escape chars used in movie-maker file-path
  //<   &lt;
  //>   &gt;
  //&   &amp;
  String sOut = StringReplace(sIn, "&lt;", "<", TReplaceFlags() << rfReplaceAll);
  sOut = StringReplace(sOut, "&gt;", ">", TReplaceFlags() << rfReplaceAll);
  sOut = StringReplace(sOut, "&amp;", "&", TReplaceFlags() << rfReplaceAll);
  return sOut;
}
//---------------------------------------------------------------------------
String __fastcall TFormMain::XmlEncode(String sIn)
{
  // XML escape chars used in movie-maker file-path
  //<   &lt;
  //>   &gt;
  //&   &amp;
  String sOut = StringReplace(sIn, "<", "&lt;",TReplaceFlags() << rfReplaceAll);
  sOut = StringReplace(sOut, ">", "&gt;", TReplaceFlags() << rfReplaceAll);
  sOut = StringReplace(sOut, "&", "&amp;",TReplaceFlags() << rfReplaceAll);
  return sOut;
}
//---------------------------------------------------------------------------
// browseforfolder function
// returns the folder or an empty string if no folder was selected
// hwnd = handle to parent window
// title = text in dialog
// folder = selected (default) folder
String __fastcall TFormMain::BrowseForFolder(HWND hwnd, String sTitle, String sFolder)
{
    String sRet;

    BROWSEINFO br;
    ZeroMemory(&br, sizeof(BROWSEINFO));
    br.lpfn = BrowseCallbackProc;
    br.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
    br.hwndOwner = hwnd;
    br.lpszTitle = sTitle.c_str();
    br.lParam = (LPARAM)sFolder.c_str();

    LPITEMIDLIST pidl;
    if ((pidl = SHBrowseForFolder(&br)) != NULL)
    {
        wchar_t buffer[MAX_PATH];
        if (SHGetPathFromIDList(pidl, buffer))
          sRet = String(buffer);
    }

    return sRet;
}
//---------------------------------------------------------------------------
String __fastcall TFormMain::RecurseFind(String sPath, String sFile){

  sPath =  IncludeTrailingPathDelimiter(sPath);
  wchar_t folder[MAX_PATH];
  wcscpy(folder, sPath.w_str());

  wchar_t filename[MAX_PATH];
  wcscpy(filename, sFile.w_str());

  wchar_t fullpath[MAX_PATH];
  fullpath[0] = L'\0';

  wchar_t delimiter[2] = {0};
  delimiter[0] = sPath[sPath.Length()];

  GbQuit = false;
  TimerSearchTimeout->Interval = TIMEOUT_SEARCH;
  GSearchFileName = sFile;
  TimerSearchTimeout->Enabled = true;

  int iLen = findfile_recursive(folder, filename, fullpath, delimiter);

  if (TimerSearchTimeout->Enabled)
    TimerSearchTimeout->Enabled = false;

  if (!iLen || GbQuit)
    return "";

  return String(fullpath);
}
//---------------------------------------------------------------------------
int __fastcall TFormMain::findfile_recursive(const wchar_t *folder,
                  const wchar_t *filename, wchar_t *fullpath, wchar_t *delimiter)
{
    wchar_t wildcard[MAX_PATH];
    int lenFolder = wcslen(folder);
    if (!lenFolder || folder[lenFolder-1] != delimiter[0])
      swprintf(wildcard, L"%s%c*", folder, delimiter[0]);
    else
      swprintf(wildcard, L"%s*", folder);
    WIN32_FIND_DATA fd;
    HANDLE handle = FindFirstFile(wildcard, &fd);
    if(handle == INVALID_HANDLE_VALUE) return 0;
    do
    {
      Application->ProcessMessages(); // allows GbQuit to be set in timer-timeout hook

      if (GbQuit)
        return 0;

      if(wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0)
        continue;

      wchar_t path[MAX_PATH];
      if (!lenFolder || folder[lenFolder-1] != delimiter[0])
        swprintf(path, L"%s%c%s", folder, delimiter[0], fd.cFileName);
      else
        swprintf(path, L"%s%s", folder, fd.cFileName);

      if(_wcsicmp(fd.cFileName, filename) == 0)
        wcscpy(fullpath, path);
      else if(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
        findfile_recursive(path, filename, fullpath, delimiter);
      if(wcslen(fullpath))
        break;
    } while(!GbQuit && FindNextFile(handle, &fd));
    FindClose(handle);
    return wcslen(fullpath);
}
//---------------------------------------------------------------------------
//String __fastcall TFormMain::GetSpecialFolder(int csidl)
//{
//  HMODULE h = NULL;
//  String sOut;
//  WideChar* buf = NULL;
//
//  try{
//    h = LoadLibraryW(L"Shell32.dll");
//
//    if (h != NULL){
//      tGetFolderPath pGetFolderPath;
//      pGetFolderPath = (tGetFolderPath)GetProcAddress(h, "SHGetFolderPathW");
//
//      if (pGetFolderPath != NULL){
//        buf = new WideChar[MAX_PATH];
//        buf[0] = L'\0';
//
//        if ((*pGetFolderPath)(Application->Handle, csidl, NULL, SHGFP_TYPE_CURRENT, (LPTSTR)buf) == S_OK)
//          sOut = String(buf);
//      }
//    }
//
//  }
//  __finally{
//    if (h != NULL) FreeLibrary(h);
//    if (buf != NULL) delete [] buf;
//  }
//
//  return sOut;
//}
//---------------------------------------------------------------------------

