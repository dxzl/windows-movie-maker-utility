//---------------------------------------------------------------------------
// Copyright 2008 Scott Swift - This program is distributed under the
// terms of the GNU General Public License.
//---------------------------------------------------------------------------
// DESCRIPTION:
// This program converts a servers.ini file for the mIRC IRC chat-client
// into a server-list the XiRCON IRC chat-client can use and saves them
// in your registry. The next time you start XiRCON, the new servers will
// be appended to your old servers-list. The program also weeds-out
// redundant servers.
//
// Download the mIRC servers list:
//   http://mIRC.co.uk/servers.ini
//
// For instructions, run ServerImporter.exe and click the Help button.
//---------------------------------------------------------------------------
#ifndef MainFormH
#define MainFormH
//---------------------------------------------------------------------------
//#include <Vcl.OleCtrls.hpp>
//#include <Vcl.DdeMan.hpp>
#include <Vcl.ComCtrls.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.FileCtrl.hpp>
#include <Vcl.Forms.hpp>
#include <Vcl.Dialogs.hpp>
#include <Vcl.ExtCtrls.hpp>
#include <Vcl.Mask.hpp>
#include <Vcl.Buttons.hpp>
#include <Vcl.Menus.hpp>
#include <System.Classes.hpp>
#include <System.IOUtils.hpp>
#include <Registry.hpp>
#include <stdio.h>
#include <Windows.h>
//---------------------------------------------------------------------------
class TFormMain : public TForm
{
__published:  // IDE-managed Components
  TPanel *Panel1;
  TButton *ButtonReadProjectFile;
  TOpenDialog *OpenDialog1;
  TMemo *Memo1;
  TSaveDialog *SaveDialog1;
  TGroupBox *GroupBox1;
  TLabel *LabelProjectFilePath;
  TMainMenu *MainMenu1;
  TMenuItem *ools1;
  TMenuItem *CopyMovieImageFilesToNewFolder1;
  TProgressBar *ProgressBar1;
  TMenuItem *Moveprojecttonewfolder1;
  TLabel *LabelMediaFolderPath;
  TButton *ButtonSelectMediaFolder;
  TButton *ButtonHelp;
  TTimer *TimerSearchTimeout;
  TButton *ButtonApplyNewRootPath;
  TButton *ButtonSaveFile;
  void __fastcall ButtonReadProjectFileClick(TObject *Sender);
  void __fastcall ButtonApplyNewRootPathClick(TObject *Sender);
  void __fastcall ButtonHelpClick(TObject *Sender);
  void __fastcall ButtonSaveFileClick(TObject *Sender);
  void __fastcall ButtonSelectMediaFolderClick(TObject *Sender);
  void __fastcall FormShow(TObject *Sender);
  void __fastcall CopyMovieImageFilesToNewFolder1Click(TObject *Sender);
  void __fastcall Moveprojecttonewfolder1Click(TObject *Sender);
  void __fastcall TimerSearchTimeoutTimer(TObject *Sender);
protected:
  void __fastcall WMDropFile(TWMDropFiles &Msg);

BEGIN_MESSAGE_MAP
    //add message handler for WM_DROPFILES
    // NOTE: All of the TWM* types are in messages.hpp!
    VCL_MESSAGE_HANDLER(WM_DROPFILES, TWMDropFiles, WMDropFile)
END_MESSAGE_MAP(TComponent)

private:  // User declarations
//  int __fastcall findfile_recursive(const wchar_t *folder_path,
//       const wchar_t *filename, wchar_t *found_full_path, wchar_t *delimiter);
//  String __fastcall RecurseFind(String sPath, String sFile);

  int __fastcall findfile_recursive(const wchar_t *folder,
                const wchar_t *filename, wchar_t *fullpath, wchar_t *delimiter);
  String __fastcall RecurseFind(String sPath, String sFile);


  String __fastcall BrowseForFolder(HWND hwnd, String sTitle, String sFolder);
//  bool __fastcall GetListsOfFilesDirsInCurrentDir(TStringList *slFiles, TStringList *slDirs);
  String __fastcall CommonPath(TStringList *slPaths);
  String __fastcall GetCommonPath();
  String __fastcall GetSpecialFolder(int csidl);
  void __fastcall LoadFile(void);
  void __fastcall AddFilesToStringList(TStringList* slFiles);
  void __fastcall RecurseFileAdd(TStringList* slFiles);
  String __fastcall SelectFolder(String sCaption, String sPath);
  String __fastcall XmlEncode(String sIn);
  String __fastcall XmlDecode(String sIn);
  void __fastcall CopyOrMoveProject(bool bMove);
  int __fastcall GetFilePathCount();

  String GProjectFileName, GMediaFolderPath, GCommonPath, GSearchFileName;
  bool GbQuit, GbQuitAll;

  public:    // User declarations
  __fastcall TFormMain(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormMain *FormMain;

typedef HRESULT __stdcall (*tGetFolderPath)(HWND hwndOwner, int nFolder, HANDLE hToken, DWORD dwFlags, LPTSTR pszPath);
//---------------------------------------------------------------------------
#endif
