using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using ShellLib;


namespace ThunderMain.Shell {

    /// <summary>
    /// Column Handler for all files - adds an MD5 hash of the file as a new column
    /// 
    /// Needs to be registered with RegAsm.exe and installed in the gac with GacUtil.exe
    /// https://www.codeproject.com/Articles/3747/Explorer-column-handler-shell-extension-in-C
    /// </summary>
    [Guid("875b8b0a-48c3-491a-adf7-4ede573a2b95"), ComVisible(true)]
    public class MD5ColumnHandler : ColumnProvider {
        public static String ColumnId = "c-id-test";
        static Guid GUID_A = Guid.Parse("cb099851-9766-4ac9-94bd-f2a4a8b07c7e");
        static Guid GUID_B = Guid.Parse("471581fb-381b-4a11-96fb-7d53e29755a8");
        const int S_OK = 0;
        const int S_FALSE = 1;

        public MD5ColumnHandler() {
        }

        public override int Initialize(LPCSHCOLUMNINIT psci) {
            return S_OK;
        }
        public override int GetColumnInfo(int dwIndex, out SHCOLUMNINFO psci) {
            psci = new SHCOLUMNINFO();

            if (dwIndex == 0) {
                try {
                    psci.scid.fmtid = GUID_A;
                    psci.scid.pid = 0;

                    // Cast to a ushort, because a VARTYPE is ushort and a VARENUM is int
                    psci.vt = (ushort)VarEnum.VT_BSTR;
                    psci.fmt = LVCFMT.LEFT;
                    psci.cChars = 24;

                    psci.csFlags = SHCOLSTATE.TYPE_STR;
                    psci.wszTitle = "LocalizedResourceName";
                    psci.wszDescription = "説明文 どこで表示されるの？";
                } catch (Exception e) {
                    MessageBox.Show(e.Message);
                    return S_FALSE;
                }

            } else if (dwIndex == 1) {
                try {
                    psci.scid.fmtid = GUID_B;
                    psci.scid.pid = 0;

                    // Cast to a ushort, because a VARTYPE is ushort and a VARENUM is int
                    psci.vt = (ushort)VarEnum.VT_BSTR;
                    psci.fmt = LVCFMT.LEFT;
                    psci.cChars = 40;

                    psci.csFlags = SHCOLSTATE.TYPE_STR;
                    psci.wszTitle = "カラム名B";
                    psci.wszDescription = "説明文 どこで表示されるの？";
                } catch (Exception e) {
                    MessageBox.Show(e.Message);
                    return S_FALSE;
                }

            } else {
                // 0と1ではOK返してるけど、2以降ではfalseを返す
                return S_FALSE;
            }
            return S_OK;
        }
        public override int GetItemData(LPCSHCOLUMNID pscid, LPCSHCOLUMNDATA pscd, out object pvarData) {
            pvarData = string.Empty;

            if (pscid.fmtid == GUID_B) {
                pvarData = "GUID_B";
                return S_OK;
            }
            var fileAttributes = (FileAttributes)pscd.dwFileAttributes;
            var filePath = pscd.wszFile;
            // ディレクトリオンリー
            if ((fileAttributes & FileAttributes.Directory) == 0) {
                return S_FALSE;
            }

            // Only service known columns
            if (pscid.fmtid != GUID_A || pscid.pid != 0) {
                return S_FALSE;
            }
            try {
                // ファイルのdesktop.iniを習得する
                var desktopIniPath = System.IO.Path.Combine(filePath, "desktop.ini");
                if (System.IO.File.Exists(desktopIniPath) == false) {
                    return S_FALSE;
                }
                String resourceName = "";
                using (FileStream stream = new FileStream(desktopIniPath, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                    using (StreamReader sr = new StreamReader(stream, Encoding.GetEncoding("Shift_JIS"))) {
                        String line;
                        System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(@"LocalizedResourceName\s*=(.+)");
                        while ((line = sr.ReadLine()) != null) {
                            var match = reg.Match(line);
                            if (match.Success) {
                                resourceName = match.Groups[1].ToString().Trim();
                            }
                        }
                    }
                }
                pvarData = resourceName;
            } catch (UnauthorizedAccessException e) {
                MessageBox.Show(e.Message);
                return S_FALSE;
            } catch (Exception e) {
                MessageBox.Show(e.Message);
                return S_FALSE;
            }
            return S_OK;
        }
    }
}
