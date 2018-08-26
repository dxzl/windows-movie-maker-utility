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
#include <Registry.hpp>
//---------------------------------------------------------------------------
// File path to MovieMaker Prokects
#define DEFAULT_PATH (GetSpecialFolder(CSIDL_MYDOCUMENTS) + "\\MovieMaker\\MyProject\\")
//---------------------------------------------------------------------------
class TFormMain : public TForm
{
__published:  // IDE-managed Components
  TPanel *Panel1;
  TButton *ButtonReadProjectFile;
  TOpenDialog *OpenDialog1;
  TMemo *Memo1;
  TButton *ButtonApplyNewRootPath;
  TButton *ButtonAbout;
  TButton *ButtonSaveFile;
  TSaveDialog *SaveDialog1;
  TEdit *Edit1;
  TLabel *Label3;
  TGroupBox *GroupBox1;
  TLabel *LabelOldPath;
  TButton *ButtonSelectRootFolder;
  TTimer *Timer1;
  void __fastcall ButtonReadProjectFileClick(TObject *Sender);
  void __fastcall ButtonApplyNewRootPathClick(TObject *Sender);
  void __fastcall ButtonAboutClick(TObject *Sender);
  void __fastcall Edit1Change(TObject *Sender);
  void __fastcall ButtonSaveFileClick(TObject *Sender);
  void __fastcall ButtonSelectRootFolderClick(TObject *Sender);
  void __fastcall Timer1FileDropTimeout(TObject *Sender);
  void __fastcall FormShow(TObject *Sender);
protected:
  void __fastcall WMDropFile(TWMDropFiles &Msg);

BEGIN_MESSAGE_MAP
    //add message handler for WM_DROPFILES
    // NOTE: All of the TWM* types are in messages.hpp!
    VCL_MESSAGE_HANDLER(WM_DROPFILES, TWMDropFiles, WMDropFile)
END_MESSAGE_MAP(TComponent)

private:  // User declarations
  String __fastcall CommonPath(TStringList *slPaths);
  String __fastcall GetCommonPath();
  String __fastcall GetSpecialFolder(int csidl);
  void __fastcall LoadFile(void);

  bool bServersProcessed, GbIsDirectory;
  String GFileName, GNewPath, GOldCommonPath, GDragDropPath;

public:    // User declarations
  __fastcall TFormMain(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormMain *FormMain;

typedef HRESULT __stdcall (*tGetFolderPath)(HWND hwndOwner, int nFolder, HANDLE hToken, DWORD dwFlags, LPTSTR pszPath);
//---------------------------------------------------------------------------
#endif
