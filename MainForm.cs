// Copyright (C) 2003, 2004, 2011 Daisuke Arai <darai@users.sourceforge.jp>
// Copyright (C) 2005, 2007, 2008, 2010, 2013, 2014 panacoran <panacoran@users.sourceforge.jp>
// 
// This program is part of Protra.
//
// Protra is free software: you can redistribute it and/or modify it
// under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program; if not, see <http://www.gnu.org/licenses/>.
// 
// $Id: MainForm.cs 503 2014-02-10 06:10:07Z panacoran $

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using System.IO;
using System.Text;
using Protra.Lib;
using Protra.Lib.Config;
using Protra.Lib.Data;
using Protra.Lib.Dialogs;
using Protra.Lib.Lang.Builtins;
using PtSim.Dialogs;

namespace PtSim
{
    /// <summary>
    /// アプリケーションのメインフォーム
    /// </summary>
    public partial class MainForm : Form
    {
        private string _name;
        private DateTime _startTime;

        /// <summary>
        /// アプリケーションのメインエントリポイント
        /// </summary>
        [STAThread]
        private static void Main()
        {
            if (Win32API.ProcessAlreadyExists())
                return;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public MainForm()
        {
            InitializeComponent();

            // 設定ファイルの読み込み。
            (GlobalEnv.PtSimConfig = new PtSimConfig()).Load();
            (GlobalEnv.UpdateConfig = new UpdateConfig()).Load();
            (GlobalEnv.BrandData = new BrandData()).Load();
            (GlobalEnv.BrandListConfig = new BrandListConfig()).Load();

            var config = GlobalEnv.PtSimConfig;
            if (config.TimeFrame == TimeFrame.Weekly)
                radioButtonWeekly.Checked = true;
            BrandListInit();
            SelectBrandList(config.BrandListName);
            GlobalEnv.BrandListConfig.BrandListInit = BrandListInit;
            dateTimePickerHistoryFrom.Value = DateTime.Now.AddMonths(-1); //履歴の開始を1ヶ月前に設定する。
            ptFileTreeView.RootDirectory = Global.DirSystem; // システム一覧を更新する。
            ptFileTreeView.SelectedFile = config.SystemFile;
            _startTime = DateTime.MaxValue;
        }

        private void BrandListInit()
        {
            var selectedList = (BrandList)comboBoxBrandList.SelectedItem;
            var name = selectedList == null ? null : selectedList.Name;
            comboBoxBrandList.Items.Clear();
            foreach (var bl in GlobalEnv.BrandListConfig.List)
                comboBoxBrandList.Items.Add(bl);
            SelectBrandList(name);
        }

        private void SelectBrandList(string name)
        {
            foreach (BrandList bl in comboBoxBrandList.Items)
            {
                if (bl.Name != name)
                    continue;
                comboBoxBrandList.SelectedItem = bl;
                return;
            }
            comboBoxBrandList.SelectedIndex = 0; // 株価指数を選択する。
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            var config = GlobalEnv.PtSimConfig;
            if (config.Size.Height == 0) // 設定ファイルが存在しない。
                return;
            Size = config.Size;
            if (Utils.IsShowOnScreen(new Rectangle(config.Location, Size)))
                Location = config.Location;
            if (config.Maximized)
                WindowState = FormWindowState.Maximized;
            splitContainerVirtical.SplitterDistance = config.VirticalSplitterDistance;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            ConfigWrite();
        }

        private void ConfigWrite()
        {
            var config = GlobalEnv.PtSimConfig;
            var maximized = WindowState == FormWindowState.Maximized;
            var bounds = (WindowState == FormWindowState.Normal) ? Bounds : RestoreBounds;
            config.Maximized = maximized;
            config.Size = bounds.Size;
            config.Location = bounds.Location;
            config.VirticalSplitterDistance = splitContainerVirtical.SplitterDistance;
            var brandList = (BrandList)comboBoxBrandList.SelectedItem;
            config.BrandListName = (brandList == null) ? null : brandList.Name;
            config.TimeFrame = TimeFrame;
            config.SystemFile = ptFileTreeView.SelectedFile;
            config.Save();
        }

        private void deleteLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var system = ptFileTreeView.SelectedFile;
            if (system == null)
                return;
            var msg = string.Format("ファイル: {0}\n株価データ: {1}\nのシステムの実行履歴を削除します。よろしいですか？",
                                    system, TimeFrameName);
            using (new CenteredDialogHelper())
                if (MessageBox.Show(this, msg, "確認",
                                    MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation,
                                    MessageBoxDefaultButton.Button1) == DialogResult.Cancel)
                    return;
            LogData.Delete(system, TimeFrame);
            GlobalData.Delete(system, TimeFrame);
        }

        private string TimeFrameName
        {
            get { return TimeFrame == TimeFrame.Daily ? "日足" : "週足"; }
        }

        private TimeFrame TimeFrame
        {
            get { return radioButtonDaily.Checked ? TimeFrame.Daily : TimeFrame.Weekly; }
        }

        private void deleteLogAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (new CenteredDialogHelper())
                if (MessageBox.Show(this, "すべてのシステムの実行履歴を削除します。よろしいですか？", "確認", MessageBoxButtons.OKCancel,
                                    MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2) == DialogResult.Cancel)
                    return;
            LogData.DeleteAll();
            GlobalData.DeleteAll();
        }

        private void brandListEditToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var dialog = new EditBrandListDialog())
                dialog.ShowDialog(this);
            BrandListInit();
        }

        private void manualToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(Global.PathMan);
        }

        private void versionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var dialog = new VersionDialog())
                dialog.ShowDialog(this);
        }

        private void ptFileTreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            comboBoxBrandList_SelectedIndexChanged(sender, e);
        }

        [Flags]
        private enum EnableFlags
        {
            Execute = 1,
            Result = 2,
            History = 4,
            DeleteLog = 8,
            DeleteLogAll = 0x10,
            EditBrandList = 0x20,
            All = 0x3f
        }

        private void SetEnabled(EnableFlags flags)
        {
            buttonExecute.Enabled = (flags & EnableFlags.Execute) == EnableFlags.Execute;
            buttonPerformance.Enabled = (flags & EnableFlags.Result) == EnableFlags.Result;
            buttonHistory.Enabled = (flags & EnableFlags.History) == EnableFlags.History;
            deleteLogToolStripMenuItem.Enabled = (flags & EnableFlags.DeleteLog) == EnableFlags.DeleteLog;
            deleteLogAllToolStripMenuItem.Enabled = (flags & EnableFlags.DeleteLogAll) == EnableFlags.DeleteLogAll;
            brandListEditToolStripMenuItem.Enabled = (flags & EnableFlags.EditBrandList) == EnableFlags.EditBrandList;
        }

        private void comboBoxBrandList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ptFileTreeView.SelectedFile != null && comboBoxBrandList.SelectedItem != null)
            {
                if (!backgroundWorkerExecute.IsBusy && !backgroundWorkerPerformance.IsBusy)
                    SetEnabled(EnableFlags.All);
            }
            else
                SetEnabled(EnableFlags.DeleteLogAll | EnableFlags.EditBrandList);
        }

        private void buttonExecute_Click(object sender, EventArgs e)
        {
            if (backgroundWorkerExecute.IsBusy)
            {
                backgroundWorkerExecute.CancelAsync();
                Cursor = Cursors.WaitCursor;
                buttonExecute.Enabled = false;
                return;
            }
            var system = ptFileTreeView.SelectedFile;
            var brandList = comboBoxBrandList.SelectedItem;
            if (system == null || brandList == null)
                return;
            _startTime = DateTime.Now;
            buttonExecute.Text = "中断";
            SetEnabled(EnableFlags.Execute);
            textBoxExecute.Clear();
            backgroundWorkerExecute.RunWorkerAsync(new[] {system, brandList, TimeFrame});
        }

        private void backgroundWorkerExecute_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = (BackgroundWorker)sender;
            var args = (Object[])e.Argument;
            _name = (string)args[0];
            var executor = new SystemExecutor((string)args[0], (BrandList)args[1], (TimeFrame)args[2]);
            executor.Execute(worker, WrapInvoke(textBoxExecute.AppendText));
            if (worker.CancellationPending)
                e.Cancel = true;
        }

        private IAsyncResult _asyncResult;

        private AppendTextDelegate WrapInvoke(AppendTextDelegate appendText)
        {
            return text =>
            {
                if (_asyncResult != null)
                    EndInvoke(_asyncResult);
                _asyncResult = BeginInvoke(appendText, text);
            };
        }

        private void backgroundWorkerExecute_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBarExecute.Value = e.ProgressPercentage;
        }

        private void backgroundWorkerExecute_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            var config = GlobalEnv.PtSimConfig;
            if (e.Error != null)
                textBoxExecute.AppendText(e.Error.Message + "\r\nエラーが発生したので実行を中断します。");
            else if (e.Cancelled)
                textBoxExecute.AppendText("中断されました。");
            else
                textBoxExecute.AppendText("正常に終了しました。");
            if (_asyncResult != null)
            {
                EndInvoke(_asyncResult); // 最後のBeginInvokeの後始末をする。
                _asyncResult = null;
            }
            {
                // ファイルに書き出しをする
                DateTime _endTime = DateTime.Now;
                TimeSpan ts = _endTime - _startTime;
                File.WriteAllText(@".\data\log\lasttrading.txt",
                         "実行時間 = " + (int)ts.Days + "日" +
                                         (int)ts.Hours + "時間" +
                                         (int)ts.Minutes + "分" +
                                         (int)ts.Seconds + "秒\r\n" +
                         textBoxExecute.Text);
                String filePath = @".\data\log\trading_" + _name + "_" + _endTime.ToString("yyyyMMddHHmmss") + ".txt";
                File.WriteAllText(filePath, 
                         "実行時間 = " + (int)ts.Days + "日" +
                                         (int)ts.Hours + "時間" +
                                         (int)ts.Minutes + "分" +
                                         (int)ts.Seconds + "秒\r\n" +
                         textBoxExecute.Text);
            }
            Cursor = Cursors.Arrow;
            buttonExecute.Text = "実行";
            progressBarExecute.Value = 0;
            if (ptFileTreeView.SelectedFile != null && comboBoxBrandList.SelectedItem != null)
                SetEnabled(EnableFlags.All);
            else
                SetEnabled(EnableFlags.DeleteLogAll | EnableFlags.EditBrandList);
            if (config.Autoclose)
            {
               Application.Exit();
            }
        }

        private void buttonPerformance_Click(object sender, EventArgs e)
        {
            if (backgroundWorkerPerformance.IsBusy)
            {
                backgroundWorkerPerformance.CancelAsync();
                return;
            }
            var system = ptFileTreeView.SelectedFile;
            var brandList = comboBoxBrandList.SelectedItem;
            if (system == null || brandList == null)
                return;
            buttonPerformance.Text = "中断";
            SetEnabled(EnableFlags.Result | EnableFlags.History);
            richTextBoxPerformance.Clear();
            profitGraphBox.ProfitList = null;
            backgroundWorkerPerformance.RunWorkerAsync(new[] {system, brandList, TimeFrame});
        }

        private void backgroundWorkerPerformance_DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = (BackgroundWorker)sender;
            var args = (Object[])e.Argument;
            _name = (string)args[0];
            var performance = new Performance((string)args[0], (BrandList)args[1], (TimeFrame)args[2]);
            var profits = performance.Calculate(worker, WrapInvoke(richTextBoxPerformance.AppendText));
            if (worker.CancellationPending)
            {
                e.Cancel = true;
                return;
            }
            profitGraphBox.ProfitList = profits;
        }

        private void backgroundWorkerPerformance_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBarPerformance.Value = e.ProgressPercentage;
        }

        private void backgroundWorkerPerformance_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
                richTextBoxPerformance.AppendText(e.Error.Message + "\r\nエラーが発生したので計算を中断します。");
            else if (e.Cancelled)
                richTextBoxPerformance.AppendText("中断されました。");
            buttonPerformance.Text = "計算";
            progressBarPerformance.Value = 0;
            if (ptFileTreeView.SelectedFile != null && comboBoxBrandList.SelectedItem != null)
                SetEnabled(EnableFlags.All);
            else
                SetEnabled(EnableFlags.DeleteLogAll | EnableFlags.EditBrandList);
        }

        private void buttonHistory_Click(object sender, EventArgs e)
        {
            var system = ptFileTreeView.SelectedFile;
            if (system == null)
                return;
            var brands = (BrandList)comboBoxBrandList.SelectedItem;
            if (brands == null)
                return;
            List<Log> logs;
            using (var logData = new LogData(system, TimeFrame))
                logs = logData.GetLog(dateTimePickerHistoryFrom.Value, dateTimePickerHistoryTo.Value);
            var brandIdTable = new Dictionary<string, Object>();
            foreach (var id in brands.List)
                brandIdTable.Add(id, null);

            listViewHistory.Items.Clear();
            listViewHistory.BeginUpdate();
            foreach (var log in logs)
            {
                if (brandIdTable.ContainsKey(log.Code))
                {
                    var brand = GlobalEnv.BrandData[log.Code];
                    var listViewItem = new ListViewItem(new[]
                    {
                        log.Date.ToString("yy/MM/dd"),
                        brand.Code.ToString(CultureInfo.InvariantCulture),
                        brand.Name,
                        log.Price.ToString(CultureInfo.InvariantCulture),
                        log.Quantity.ToString(CultureInfo.InvariantCulture),
                        log.Order == 0 ? "買" : "売"
                    }) {BackColor = log.Order == 0 ? Color.LightBlue : Color.LightPink};
                    listViewHistory.Items.Add(listViewItem);
                }
            }
            listViewHistory.EndUpdate();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var rows = new string[listViewHistory.SelectedItems.Count];
            var i = 0;
            foreach (ListViewItem item in listViewHistory.SelectedItems)
            {
                var columns = new string[item.SubItems.Count];
                var j = 0;
                foreach (ListViewItem.ListViewSubItem subItem in item.SubItems)
                    columns[j++] = subItem.Text;
                rows[i++] = string.Join(",", columns) + "\n";
            }
            if (i == 0)
                return;
            Clipboard.SetDataObject(string.Concat(rows));
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            var config = GlobalEnv.PtSimConfig;
            if (config.Autoclose)
            {
                if (backgroundWorkerExecute.IsBusy)
                {
                    backgroundWorkerExecute.CancelAsync();
                    Cursor = Cursors.WaitCursor;
                    buttonExecute.Enabled = false;
                    return;
                }
                var system = ptFileTreeView.SelectedFile;
                var brandList = comboBoxBrandList.SelectedItem;
                if (system == null || brandList == null)
                    return;
                buttonExecute.Text = "中断";
                SetEnabled(EnableFlags.Execute);
                textBoxExecute.Clear();
                backgroundWorkerExecute.RunWorkerAsync(new[] { system, brandList, TimeFrame });
            }
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}