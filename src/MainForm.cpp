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

TFormMain *FormMain;

//---------------------------------------------------------------------------
__fastcall TFormMain::TFormMain(TComponent* Owner)
  : TForm(Owner)
{
  GProjectFilePath = DEFAULT_PATH;
  Edit1->Text = GProjectFilePath;
  GProjectFileName = "save1.txt";

  //enable drag&drop files
  ::DragAcceptFiles(this->Handle, true);
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::FormShow(TObject *Sender)
{
  String s = "FIRST: Make a backup of your .wlmp Windows Movie Maker project file!\n\n"
    "Follow the steps on the buttons, 1,2,3,4.\n\n"
    "(Hint: For steps 1  and 2 you can also drag-drop, first your project-file,\n"
    "then the folder with your movie and image files to this window. Always drag\n"
    "the project file first then the folder that has your movie and image files. Finally,\n"
    "press \"Apply new root-project path\", do any necessary manual editing in\n"
    "the window and press \"Save file\")\n\n"
    "Cheers, Scott Swift (dxzl@live.com)";
  Memo1->Lines->Text = s;
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::ButtonApplyNewRootPathClick(TObject *Sender)
{
  if (GProjectFilePath.IsEmpty())
    return;

  // add trailing delimeter in GProjectFilePath
  GProjectFilePath = IncludeTrailingPathDelimiter(GProjectFilePath);

  try
  {
    GOldCommonPath = GetCommonPath(); // get again (in case edit text manually changed)
    LabelOldPath->Caption = "Old Path: \"" + GOldCommonPath + "\""; // display it...
    int lenOldCommonPath = GOldCommonPath.Length();

    int iCount = Memo1->Lines->Count;

    for(int ii = 0 ; ii < iCount ; ii++)
    {
      String OldPath;
      String OldStr = Memo1->Lines->Strings[ii];
      int Pos1 = OldStr.Pos("filePath=\"");
      if (Pos1 == 0)
        continue;

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
      // So if there is a '\', remove text up to and including the first one
      int Pos2 = sRemainingPath.Pos("\\");
      if (Pos2 != 0)
        if (!DirectoryExists(ExtractFilePath(GProjectFilePath + sRemainingPath)))
          sRemainingPath = sRemainingPath.SubString(Pos2+1, sRemainingPath.Length()-Pos2);

      String sTempPath = GProjectFilePath;

      if (!FileExists(GProjectFilePath + sRemainingPath))
      {
        // ShowMessage("file missing: " + GProjectFilePath + sRemainingPath);

        // if we can't find the file in the new location drag-dropped folder...
        // maybe it's in the root level...
        int len = GProjectFilePath.Length();
        if (len > 0 && GProjectFilePath[len] == '\\')
        {
          int jj;
          int iEnd = 0;
          for (jj = len-1; jj > 0; jj--)
          {
            if (GProjectFilePath[jj] == '\\')
            {
              iEnd = jj;
              break;
            }
          }

          // quick and dirty attempt to just find the file one level up from the
          // drag-dropped folder
          if (iEnd > 0)
          {
            sTempPath = GProjectFilePath.SubString(1, iEnd);

            //if (FileExists(GProjectFilePath + sRemainingPath))
            //  ShowMessage("file now found at: " + GProjectFilePath + sRemainingPath);
          }
        }
      }

      // replace oldpath with newpath
      String NewStr = OldStr.SubString(1, Pos1+10-1) + XmlEncode(sTempPath + sRemainingPath);

      // add the rest of the original line
      for (; jj <= len; jj++)
        NewStr += OldStr[jj];

      Memo1->Lines->Strings[ii] = NewStr;
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
//---------------------------------------------------------------------------
String __fastcall TFormMain::GetCommonPath()
{
  TStringList *slPaths = NULL;
  String sOldCommonPath;

  try
  {
    slPaths = new TStringList();

    int iCount = Memo1->Lines->Count;

    // First, get all the old paths into a stringlist
    for(int ii = 0 ; ii < iCount ; ii++)
    {
      String OldPath;
      String OldStr = Memo1->Lines->Strings[ii];
      int Pos1 = OldStr.Pos("filePath=\"");
      if (Pos1 == 0)
        continue;

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

      slPaths->Add(ExtractFilePath(OldPath));
    }

    sOldCommonPath = CommonPath(slPaths);
  }
  __finally
  {
    if (slPaths) delete slPaths;
  }
  return sOldCommonPath;
}
//---------------------------------------------------------------------------
String __fastcall TFormMain::GetSpecialFolder(int csidl)
{
  HMODULE h = NULL;
  String sOut;
  WideChar* buf = NULL;

  try
  {
    h = LoadLibraryW(L"Shell32.dll");

    if (h != NULL)
    {
      tGetFolderPath pGetFolderPath;
      pGetFolderPath = (tGetFolderPath)GetProcAddress(h, "SHGetFolderPathW");

      if (pGetFolderPath != NULL)
      {
        buf = new WideChar[MAX_PATH];
        buf[0] = L'\0';

        if ((*pGetFolderPath)(Application->Handle, csidl, NULL, SHGFP_TYPE_CURRENT, (LPTSTR)buf) == S_OK)
          sOut = String(buf);
      }
    }

  }
  __finally
  {
    if (h != NULL) FreeLibrary(h);
    if (buf != NULL) delete [] buf;
  }

  return sOut;
}
//---------------------------------------------------------------------------
String __fastcall TFormMain::CommonPath(TStringList *slPaths)
{
    String sPath;

    int count = slPaths->Count;
    if (count < 2) return ""; // need two strings to compare

    // get "min" path length
    int min = 10000;
    for (int ii=0; ii < count; ii++)
    {
      int i = slPaths->Strings[ii].Length();
      if (i < min)
        min = i;
    }

    // compare up to "min" chars for every path
    for (int jj=1; jj <= min; jj++)
    {
      Char c1 = slPaths->Strings[0][jj];
      for (int ii=1; ii < count; ii++)
      {
        Char c2 = slPaths->Strings[ii][jj];
        if (c1 != c2)
          return sPath;
      }
      sPath += c1;
    }
    return sPath;
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::Edit1Change(TObject *Sender)
{
  GProjectFilePath = Edit1->Text;
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::ButtonAboutClick(TObject *Sender)
{
  String sHelp = "Windows Movie Maker project-file path corrector\n"
                                      "by Scott Swift 2024\n\n";
  sHelp += "The Problem: Windows Movie Maker project files use absolute file-paths.\n"
      "That means that after you make a movie, you can't change the location "
      "of either your project-file or of the media files in your movie.\n"
      "This app solves that problem!\n\n"
      "Additionally, \"Tools->Copy project to new folder\" allows you to freely make your movie "
      "with media files located anywhere on your computer then later consolidate/copy "
      "only the files used by your movie into a new folder.";
  ShowMessage(sHelp);
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::ButtonSelectRootFolderClick(TObject *Sender)
{
    Edit1->Text = SelectFolder("Select root movie-maker project directory", GProjectFilePath);
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

  if (!OpenDialog1->Execute())
  {
    Memo1->SetFocus();
    return; // Cancel
  }

  GProjectFileName = OpenDialog1->FileName;

  LoadFile();
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::ButtonSaveFileClick(TObject *Sender)
{
  try
  {
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

    if (SaveDialog1->Execute())
    {
      GProjectFileName = SaveDialog1->FileName;
      Memo1->Lines->SaveToFile(GProjectFileName, TEncoding::UTF8);
    }
  }
  catch(...)
  {
    ShowMessage("Can't save file: \"" + SaveDialog1->FileName + "\"");
  }

  Memo1->SetFocus();
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::WMDropFile(TWMDropFiles &Msg)
{
    // don't process these drag-drop files until previous ones are complete
    if (!GDragDropPath.IsEmpty())
      return;

    wchar_t* wBuf = NULL;

    try
    {
        //get dropped files count
        int cnt = ::DragQueryFileW((HDROP)Msg.Drop, -1, NULL, 0);

        if (cnt != 1) return;

        wBuf = new wchar_t[MAX_PATH];

        // Get next file-name
        if (::DragQueryFileW((HDROP)Msg.Drop, 0, wBuf, MAX_PATH) > 0)
        {
            // Load and convert file as per the file-type (either plain or rich text)
            WideString wFile(wBuf);

            // don't process this drag-drop until previous one sets m_DragDropFilePath = ""
            if (!wFile.IsEmpty())
            {
                String sFile = String(wFile); // convert to utf-8 internal string

                if (FileExists(sFile)) // is it a file? (then must be the project-file!)
                {
                    GDragDropPath = sFile;
                    GbIsDirectory = false;
                }
                else // is it a directory? (then must be the folder that has our photos and video clips!)
                {
                  if (DirectoryExists(sFile))
                  {
                    GDragDropPath = sFile + '\\';
                    GbIsDirectory = true;
                  }
                }
            }
        }

        // kick off file processing...
        if (!GDragDropPath.IsEmpty())
        {
          Timer1->Enabled = false; // stop timer (just in-case!)
          Timer1->Interval = 50;
          Timer1->OnTimer = Timer1FileDropTimeout; // set handler
          Timer1->Enabled = true; // fire event to send file
        }
    }
    __finally
    {
      try { if (wBuf != NULL) delete [] wBuf; } catch(...) {}
    }
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::Timer1FileDropTimeout(TObject *Sender)
{
    Timer1->Enabled = false;
    if (GbIsDirectory)
    {
      // handle drag-drop of the folder with video clips and photos
      GProjectFilePath = GDragDropPath;
      Edit1->Text = GProjectFilePath;
    }
    else
    {
      // handle drag-drop of the .wlmp movie project-file
      GProjectFileName = GDragDropPath;
      LoadFile();
    }

    GDragDropPath = "";
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::LoadFile(void)
{
  try
  {
    Memo1->Clear();

    // Load wlmp file
    Memo1->Lines->LoadFromFile(GProjectFileName, TEncoding::UTF8);

    GOldCommonPath = GetCommonPath();
    LabelOldPath->Caption = "Old Path: \"" + GOldCommonPath + "\""; // display it...
    GProjectFilePath = ExtractFilePath(GProjectFileName);
    Edit1->Text = GProjectFilePath;
  }
  catch(...)
  {
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
void __fastcall TFormMain::CopyOrMoveProject(bool bMove)
{
  // Copy/Move movies/images for project to new folder and generate
  // a new project file to access them.

  // 1. Load the movie-maker project file into Memo1
  Memo1->Clear();
  GProjectFileName = "";
  ButtonReadProjectFileClick(NULL);
  if (Memo1->Lines->Count == 0)
    return;

  // 2. Get destination folder for copied media files

  // GProjectFilePath must be set to the path of the project file before calling
  // SelectFolder()
  String NewMediaFilesPath =
    SelectFolder("Choose a folder where your media files will be copied!", GProjectFilePath);
//  if (!SelectDirectory(NewMediaFilesPath,
//       TSelectDirOpts() << sdAllowCreate << sdPerformCreate << sdPrompt, 0))
//    return;

  NewMediaFilesPath = IncludeTrailingPathDelimiter(NewMediaFilesPath);

  // 3. process media file names
  TStringList* pSlSourceNoExist = NULL;
  TStringList* pSlCopySuccess = NULL;

  try{
    pSlSourceNoExist = new TStringList();
    pSlCopySuccess = new TStringList();

    ProgressBar1->Position = 0;
    ProgressBar1->Step = 1;
    ProgressBar1->Min = 0;

    int iFilepathCount = 0;

    for(int ii = 0; ii < Memo1->Lines->Count; ii++)
    {
      String OldFullPath;
      String OldStr = Memo1->Lines->Strings[ii];
      int Pos1 = OldStr.Pos("filePath=\"");
      if (Pos1 != 0)
        iFilepathCount++;
    }

    ProgressBar1->Max = iFilepathCount;
    bool bAbort = false;

    for(int ii = 0; ii < Memo1->Lines->Count; ii++)
    {
      String OldFullPath;
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
        OldFullPath += c;
      }
      if (jj > len) // no trailing quote, continue
        continue;

      OldFullPath = XmlDecode(OldFullPath);

      String FileName = ExtractFileName(OldFullPath);
      String NewFullPath = NewMediaFilesPath + FileName;

      if (FileExists(OldFullPath)){
        try{
          // Note: this throws exception if file already exists!
          TFile::Copy(OldFullPath, NewFullPath, true); // overwrite true!
          pSlCopySuccess->Add(OldFullPath);
        }
        catch(...){
          ShowMessage("Aborting! Unable to copy \"" + OldFullPath +
                      "\" to \"" + NewFullPath + "\"");
          bAbort = true;
          break;
        }
        Application->ProcessMessages();
      }
      else
        pSlSourceNoExist->Add(OldFullPath);

      String sNewLine = OldStr.SubString(1, Pos1+10-1);
      sNewLine += XmlEncode(NewFullPath);
      sNewLine += OldStr.SubString(jj, OldStr.Length());
      Memo1->Lines->Strings[ii] = sNewLine;
    }

    // Move to line 0, Character 0:
    Memo1->SelStart = Memo1->Perform(EM_LINEINDEX, 0, 0);
    Memo1->SelLength = 0;
    Memo1->Perform(EM_SCROLLCARET, 0, 0);
    Memo1->SetFocus();

    ProgressBar1->Position = 0;

    if (pSlSourceNoExist && pSlCopySuccess){
      if (!bAbort){
        if (pSlCopySuccess->Count){
          String s1;

          if (bMove){
            int iFailCount = 0;
            for (int ii=0; ii < pSlCopySuccess->Count; ii++){
              try{
                TFile::Delete(pSlCopySuccess->Strings[ii]);
              }
              catch(...){
                iFailCount++;
              }
            }

            if (iFailCount){
              s1 += "Aborting. Unable to delete " + String(iFailCount) +
               " files after copying! Do you have permission to delete these files?";
              bAbort = true;
            }
            else{
              s1 += String(pSlCopySuccess->Count) + " of " +
                 String(iFilepathCount-pSlSourceNoExist->Count) +
                    " file(s) succesfuly moved!";
            }
          }
          else{
            s1 += String(pSlCopySuccess->Count) + " of " +
                 String(iFilepathCount-pSlSourceNoExist->Count) +
                    " file(s) succesfuly copied!";
          }
          ShowMessage(s1);
        }

        if (pSlSourceNoExist->Count){
          String s1 = "The project file references " + String(pSlSourceNoExist->Count) +
           " file(s) that don't exist in the specified location!\n\n";

          for (int ii=0; ii < pSlSourceNoExist->Count && ii < MAX_DISPLAY_FILES; ii++)
            s1 += pSlSourceNoExist->Strings[ii] + "\n";

          if (!bAbort)
            s1 += "\nYou can choose \"Cancel\" in the next dialog or you can save "
              "the project using a new name and give it a try!";

          ShowMessage(s1);
        }

        // before calling ButtonSaveFileClick(NULL) we expect GProjectFileName
        // to have the full project file path and name!
        if (!bAbort)
          ButtonSaveFileClick(NULL);
      }

      delete pSlSourceNoExist;
      delete pSlCopySuccess;
    }
  }
  catch(...){
    ShowMessage("Exception thrown!");
  }
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
//!!!!!!!!! TODO: I want to add the ability for the program to recurse through subfolders
// to find movie clips/photos
//void __fastcall TFormMain::RecurseFileAdd(TStringList* slFiles)
//// Uses the WideString versions of the Win32 API FindFirstFile and FindNextFile directly
//// and converts the resulting paths to UTF-8 for storage in an ordinary TStringList
////
//// Use SetCurrentDirectory() to set our root directory or TOpenDialog sets it also...
//{
//  Application->ProcessMessages();
//
//  if (slFiles == NULL) return;
//
//  TStringList* slSubDirs = new TStringList();
//  if (slSubDirs == NULL) return;
//
//  TWin32FindDataW sr;
//  HANDLE hFind = NULL;
//  TFindexInfoLevels l = FindExInfoStandard; // FindExInfoBasic was defined later!
//  TFindexSearchOps s = FindExSearchLimitToDirectories;
//
//  try
//  {
//    hFind = FindFirstFileExW(L"*", l, &sr, s, NULL, (DWORD)FIND_FIRST_EX_LARGE_FETCH);
//
//    // Get list of subdirectories into a stringlist
//    if (hFind != INVALID_HANDLE_VALUE)
//    {
//      do
//      {
//        if ((sr.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
//        {
//          int len = wcslen(sr.cFileName);
//
//          if (len == 1 && sr.cFileName[0] == L'.')
//            continue;
//          if (len == 2 && sr.cFileName[0] == L'.' && sr.cFileName[1] == L'.')
//            continue;
//
////          slSubDirs->Add(WideToUtf8(WideString(sr.cFileName)));
//          slSubDirs->Add(sr.cFileName);
//        }
//      } while (FindNextFileW(hFind, &sr) == TRUE);
//    }
//  }
//  __finally
//  {
//    try { if (hFind != NULL) FindClose(hFind); } catch(...) {}
//  }
//
//  AddFilesToStringList(slFiles);
//
//  // Get songs in all subdirectories
//  for (int ii = 0; ii < slSubDirs->Count; ii++)
//  {
//    Application->ProcessMessages();
//    if (Application->Terminated || (int)GetAsyncKeyState(VK_ESCAPE) < 0)
//      break;
//
//    if (SetCurrentDirectoryW(slSubDirs->Strings[ii].w_str()))
//    {
//      RecurseFileAdd(slFiles);
//      SetCurrentDirectoryW(L"..");
//    }
//  }
//
//  delete slSubDirs;
//}
////---------------------------------------------------------------------------
//void __fastcall TFormMain::AddFilesToStringList(TStringList* slFiles)
//// slFiles is in UTF-8!
//{
//  TWin32FindDataW sr;
//  HANDLE hFind = NULL;
//  TFindexInfoLevels l = FindExInfoStandard; // FindExInfoBasic was defined later!
//  TFindexSearchOps s = FindExSearchNameMatch;
//
//  try
//  {
//    // Get the current directory
//    String wdir = GetCurrentDir();
//
//    hFind = FindFirstFileExW(L"*", l, &sr, s, NULL, (DWORD)FIND_FIRST_EX_LARGE_FETCH);
//
//    // Get list of files into a stringlist
//    if (hFind != INVALID_HANDLE_VALUE)
//    {
//      String ws;
//
//      // Don't add these file-types...
//      int mask = FILE_ATTRIBUTE_DIRECTORY|FILE_ATTRIBUTE_SYSTEM|FILE_ATTRIBUTE_HIDDEN;
//
//      do
//      {
//        if ((sr.dwFileAttributes & mask) == 0)
//        {
//          ws = wdir + L"\\" + sr.cFileName;
//          slFiles->Add(ws);
//        }
//      } while (FindNextFileW(hFind, &sr) == TRUE);
//    }
//  }
//  __finally
//  {
//    try { if (hFind != NULL) ::FindClose(hFind); } catch(...) {}
//  }
//}
//---------------------------------------------------------------------------

