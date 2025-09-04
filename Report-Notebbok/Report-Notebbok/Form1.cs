using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading.Tasks;

namespace WinApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            this.Text = "Report-Notebook";

            Button btnRun = new Button();
            btnRun.Text = "Run Report-Notebook";
            btnRun.Width = 190;
            btnRun.Height = 30;
            btnRun.Top = 50;
            btnRun.Left = 50;
            btnRun.Click += BtnRun_Click;

            this.Controls.Add(btnRun);
        }

        private async void BtnRun_Click(object sender, EventArgs e)
        {
            // ProgressBar + Label
            ProgressBar progress = new ProgressBar();
            progress.Width = 190;
            progress.Height = 30;
            progress.Minimum = 0;
            progress.Maximum = 100;
            progress.Value = 0;
            progress.Top = 100;
            progress.Left = 50;
            this.Controls.Add(progress);

            Label lblStatus = new Label();
            lblStatus.Text = "เริ่มงาน...";
            lblStatus.Top = 130;
            lblStatus.Left = 50;
            lblStatus.Width = 300;
            this.Controls.Add(lblStatus);

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                // -----------------------------
                // Step 1: รัน Task Scheduler ตัวแรก
                // -----------------------------
                lblStatus.Text = "รัน Task Scheduler แรก...";
                progress.Value = 10;

                ProcessStartInfo psiBackup = new ProcessStartInfo
                {
                    FileName = "schtasks",
                    Arguments = @"/run /s TSEDB /u tse\administrator /p scsadmin /tn ""Report-Notebook""",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process procBackup = Process.Start(psiBackup))
                {
                    procBackup.WaitForExit();
                }

                progress.Value = 30;

                // -----------------------------
                // Step 2: เปิด Excel และ Refresh HM
                // -----------------------------
                lblStatus.Text = "เปิด Excel และ Refresh HM...";

                string workbookPathHM = @"\\172.24.3.139\Liitichai_Yorach\Script\Report-HM.xlsm";

                if (!File.Exists(workbookPathHM))
                {
                    using (OpenFileDialog ofd = new OpenFileDialog()
                    {
                        Title = "เลือกไฟล์ Report-HM.xlsm",
                        Filter = "Excel Macro-Enabled (*.xlsm)|*.xlsm|All files (*.*)|*.*",
                        CheckFileExists = true,
                        Multiselect = false
                    })
                    {
                        if (ofd.ShowDialog() != DialogResult.OK)
                        {
                            MessageBox.Show("ไม่พบไฟล์: " + workbookPathHM, "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        workbookPathHM = ofd.FileName;
                    }
                }

                excelApp = new Excel.Application { Visible = false };
                workbook = excelApp.Workbooks.Open(workbookPathHM);

                excelApp.Run("REFRESH");
                workbook.RefreshAll();
                await Task.Run(() => excelApp.CalculateUntilAsyncQueriesDone());

                workbook.Save();
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
                workbook = null;

                progress.Value = 50;
                lblStatus.Text = "HM เสร็จสิ้น";

                // -----------------------------
                // Step 2.2: เปิด Excel และ Refresh TG
                // -----------------------------
                lblStatus.Text = "เปิด Excel และ Refresh TG...";

                string workbookPathTG = @"\\172.24.3.139\Liitichai_Yorach\Script\Report-TG.xlsm";

                if (!File.Exists(workbookPathTG))
                {
                    using (OpenFileDialog ofd = new OpenFileDialog()
                    {
                        Title = "เลือกไฟล์ Report-TG.xlsm",
                        Filter = "Excel Macro-Enabled (*.xlsm)|*.xlsm|All files (*.*)|*.*",
                        CheckFileExists = true,
                        Multiselect = false
                    })
                    {
                        if (ofd.ShowDialog() != DialogResult.OK)
                        {
                            MessageBox.Show("ไม่พบไฟล์: " + workbookPathTG, "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        workbookPathTG = ofd.FileName;
                    }
                }

                workbook = excelApp.Workbooks.Open(workbookPathTG);

                excelApp.Run("REFRESH");
                workbook.RefreshAll();
                await Task.Run(() => excelApp.CalculateUntilAsyncQueriesDone());

                workbook.Save();
                workbook.Close(false);
                Marshal.ReleaseComObject(workbook);
                workbook = null;

                progress.Value = 70;
                lblStatus.Text = "TG เสร็จสิ้น";

                // -----------------------------
                // Step 3: รัน Task Scheduler ตัวสอง
                // -----------------------------
                lblStatus.Text = "รัน Task Scheduler สุดท้าย...";
                progress.Value = 80;

                ProcessStartInfo psiNew = new ProcessStartInfo
                {
                    FileName = "schtasks",
                    Arguments = @"/run /s TSEDB /u tse\administrator /p scsadmin /tn ""Report-Notebook2""",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process procNew = Process.Start(psiNew))
                {
                    procNew.WaitForExit();
                }

                progress.Value = 100;
                lblStatus.Text = "เสร็จสมบูรณ์!";

                MessageBox.Show("✅ สำเร็จทั้งหมดแล้ว", "สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("❌ เกิดข้อผิดพลาด: " + ex.Message, "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                try
                {
                    if (workbook != null)
                    {
                        workbook.Close(false);
                        Marshal.ReleaseComObject(workbook);
                    }
                }
                catch { }

                try
                {
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                }
                catch { }

                workbook = null;
                excelApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
