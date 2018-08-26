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

TFormMain *FormMain;

//---------------------------------------------------------------------------
__fastcall TFormMain::TFormMain(TComponent* Owner)
  : TForm(Owner)
{
  GNewPath = DEFAULT_PATH;
  Edit1->Text = GNewPath;
  GFileName = "save1.txt";

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
    "the project file first then the folder that has your video clips. Finally,\n"
    "press \"Apply new root-project path\", do any necessary manual editing in\n"
    "the window and press \"Save file\")\n\n"
    "Cheers, Scott Swift (dxzl@live.com)";
  Memo1->Lines->Text = s;
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::ButtonApplyNewRootPathClick(TObject *Sender)
{

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

//      String FileName = ExtractFileName(OldPath);
//      String FilePath = ExtractFilePath(OldPath);
      String sRemainingPath = OldPath.SubString(lenOldCommonPath+1, OldPath.Length());

      // check backslash char in GNewPath
      int npLen = GNewPath.Length();
      if (npLen && GNewPath[npLen] != '\\')
        GNewPath += '\\';

      // Need to remove the first subdirectory if it does not exist...
      // It's being replaced with the new folder we drag-dropped...
      // So if there is a '\', remove text up to and including the first one
      int Pos2 = sRemainingPath.Pos("\\");
      if (Pos2 != 0)
        if (!DirectoryExists(ExtractFilePath(GNewPath + sRemainingPath)))
          sRemainingPath = sRemainingPath.SubString(Pos2+1, sRemainingPath.Length()-Pos2);

      // replace oldpath with newpath
      String NewStr = OldStr.SubString(1, Pos1+10-1) + GNewPath + sRemainingPath;

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
  GNewPath = Edit1->Text;
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::ButtonAboutClick(TObject *Sender)
{
  ShowMessage("Windows Movie Maker project-file path corrector\n"
                                      "by Scott Swift 2018");
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::ButtonSelectRootFolderClick(TObject *Sender)
{
    UnicodeString selDir;
    TSelectDirExtOpts opt;

    opt.Clear();
    opt += TSelectDirExtOpts()<<sdNewFolder;
    opt += TSelectDirExtOpts()<<sdShowEdit;
    opt += TSelectDirExtOpts()<<sdNewUI;

    if (SelectDirectory( L"Select root movie-maker project directory", L"",
                                      GNewPath, opt, this ))
        Edit1->Text = GNewPath;
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

  GFileName = OpenDialog1->FileName;

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
    SaveDialog1->FileName = GFileName;

    if (SaveDialog1->Execute())
    {
      GFileName = SaveDialog1->FileName;
      Memo1->Lines->SaveToFile(GFileName);
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
      GNewPath = GDragDropPath;
      Edit1->Text = GNewPath;
    }
    else
    {
      // handle drag-drop of the .wlmp movie project-file
      GFileName = GDragDropPath;
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

    // Load ini file
    Memo1->Lines->LoadFromFile(GFileName);

    GOldCommonPath = GetCommonPath();
    LabelOldPath->Caption = "Old Path: \"" + GOldCommonPath + "\""; // display it...
    GNewPath = ExtractFilePath(GFileName);
    Edit1->Text = GNewPath;
  }
  catch(...)
  {
    ShowMessage("Can't load file: \"" + GFileName + "\"");
  }

  Memo1->SetFocus();
}
//---------------------------------------------------------------------------

