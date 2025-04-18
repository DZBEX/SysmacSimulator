using Sunny.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SysmacSimulator
{
    public partial class SysmacSimulatorTool : UIForm
    {
        private SysmacSimulatorHelper simDriver = new SysmacSimulatorHelper();
        private bool isConnected;
        private CancellationTokenSource _cancellationTokenSource;
        private List<SimulatorVariable> _variableList;
        private bool isDataLoaded; // 标志变量，表示数据是否已经加载到列表中
        private Stopwatch _resizeStopwatch;
        private int _resizeInterval = 10000; // 5 秒调整一次列宽

        public SysmacSimulatorTool()
        {
            InitializeComponent();
            _variableList = new List<SimulatorVariable>();
            uiDataGridView1.VirtualMode = true;
            uiDataGridView1.CellValueNeeded += UiDataGridView1_CellValueNeeded;
            uiDataGridView1.CellValuePushed += UiDataGridView1_CellValuePushed;
            uiDataGridView1.NewRowNeeded += UiDataGridView1_NewRowNeeded;
            uiDataGridView1.RowCount = 0;
            isDataLoaded = false; // 初始化标志变量
            _resizeStopwatch = new Stopwatch();
            _resizeStopwatch.Start();
        }

        #region 通信调试

        /// <summary>
        /// 选择全局变量文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_ChooseFolder_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Text files (*.txt)|*.txt"
            };
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                // 全局变量表
                simDriver.PopulateFromFile(filePath);

                // 清空现有的列表
                _variableList.Clear();

                // 将字典数据添加到列表
                foreach (var kvp in simDriver.variableDictionary)
                {
                    _variableList.Add(kvp.Value);
                }

                // 配置 DataGridView 列
                uiDataGridView1.Columns.Clear();
                uiDataGridView1.Columns.Add("VariableName", "Variable Name");
                uiDataGridView1.Columns.Add("RevisionString", "Revision");
                uiDataGridView1.Columns.Add("AddressString", "Address");
                uiDataGridView1.Columns.Add("Size", "Size");
                uiDataGridView1.Columns.Add("Type", "Type");
                uiDataGridView1.Columns.Add("LowIndex", "Low Index");
                uiDataGridView1.Columns.Add("HighIndex", "High Index");
                uiDataGridView1.Columns.Add("Value", "Value");

                uiDataGridView1.RowCount = _variableList.Count;
                uiDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                uiDataGridView1.Refresh();

                // 自动调整列宽
                uiDataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

                isDataLoaded = true; // 设置标志变量为 true

                // 取消当前的 ReadTask 线程
                if (_cancellationTokenSource != null)
                {
                    _cancellationTokenSource.Cancel();
                    _cancellationTokenSource.Dispose();
                }

                // 重启 ReadTask 线程
                if (isConnected)
                {
                    _cancellationTokenSource = new CancellationTokenSource();
                    ReadTask(_cancellationTokenSource.Token);
                }
            }
        }

        /// <summary>
        /// 连接SysMac
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiSymbolButton1_Click(object sender, EventArgs e)
        {
            // 连接
            if (!isConnected)
            {
                var returnedvalues = simDriver.Connect();
                isConnected = returnedvalues.Item1;
                AddLog(returnedvalues.Item1 ? 0 : 1, returnedvalues.Item2);
                uiSymbolButton1.Text = "断开连接";
                uiSymbolButton1.Symbol = 361735;
                uiLedBulb2.Blink = true;
                Btn_ChooseFolder.Enabled = true;

                if (isDataLoaded)
                {

                    // 重启 ReadTask 线程
                    if (isConnected)
                    {
                        _cancellationTokenSource = new CancellationTokenSource();
                        ReadTask(_cancellationTokenSource.Token);
                    }
                }
            }
            else
            {
                var returnedvalues = simDriver.Disconnect();
                bool isDisconnect = returnedvalues.Item1;
                AddLog(returnedvalues.Item1 ? 0 : 1, returnedvalues.Item2);
                isConnected = false;
                uiSymbolButton1.Text = "连接SysMac";
                uiSymbolButton1.Symbol = 361633;
                uiLedBulb2.Blink = false;
                Btn_ChooseFolder.Enabled = false;

                // 取消读取任务
                if (_cancellationTokenSource != null)
                {
                    _cancellationTokenSource.Cancel();
                    _cancellationTokenSource.Dispose();
                }
            }
        }

        /// <summary>
        /// 读取变量
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Read_Btn_Click(object sender, EventArgs e)
        {
            var returnedvalues = simDriver.ReadVariable(txt_variableName.Text);
            AddLog(returnedvalues.Item1 ? 0 : 1, returnedvalues.Item2);
        }

        /// <summary>
        /// 写入变量
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Write_Btn_Click(object sender, EventArgs e)
        {
            var returnedvalues = simDriver.WriteVariable(txt_variableName.Text, txt_Writevalue.Text);
            AddLog(returnedvalues.Item1 ? 0 : 1, returnedvalues.Item2);
        }

        /// <summary>
        /// 通用写入日志的方法
        /// </summary>
        /// <param name="listView"></param>
        /// <param name="index"></param>
        /// <param name="log"></param>
        public void AddLog(int index, string log)
        {
            ListViewItem listViewItem = new ListViewItem("  " + DateTime.Now.ToString("HH:mm:ss"), index);

            if (this.lst_Info.InvokeRequired)
            {
                this.lst_Info.Invoke(new Action(() =>
                {
                    listViewItem.SubItems.Add(log);
                    this.lst_Info.Items.Insert(0, listViewItem);
                }));
            }
            else
            {
                listViewItem.SubItems.Add(log);
                //保证最新的日志在最上面
                this.lst_Info.Items.Add(listViewItem);
                this.lst_Info.Items[this.lst_Info.Items.Count - 1].EnsureVisible();
            }
        }

        private async void ReadTask(CancellationToken cancellationToken)
        {
            try
            {
                while (isConnected && !cancellationToken.IsCancellationRequested)
                {
                    for (int i = 0; i < _variableList.Count; i++)
                    {
                        try
                        {
                            var variable = _variableList[i];
                            var returnedvalues = simDriver.ReadVariable(variable.VariableName);
                            if (returnedvalues.Item1)
                            {
                                // 更新变量的 Value 属性
                                variable.Value = returnedvalues.Item3;
                            }
                            else
                            {
                                AddLog(1, $"变量: {variable.VariableName}, 读取失败: {returnedvalues.Item2}");
                            }
                        }
                        catch (Exception ex)
                        {
                            AddLog(1, $"读取变量 {i} 时发生错误: {ex.Message}");
                        }
                    }

                    // 通知 DataGridView 数据已更改
                    uiDataGridView1.Invoke(new Action(() =>
                    {
                        uiDataGridView1.Invalidate();
                    }));

                    // 每 10 秒调整一次列宽
                    if (_resizeStopwatch.ElapsedMilliseconds >= _resizeInterval)
                    {
                        uiDataGridView1.Invoke(new Action(() =>
                        {
                            uiDataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                        }));
                        _resizeStopwatch.Restart();
                    }

                    // 防止线程占用过多资源，设置一个短暂的延迟
                    await Task.Delay(100, cancellationToken); // 延迟时间可以根据需要调整
                }
            }
            catch (OperationCanceledException)
            {
                // 任务被取消，无需处理
            }
            catch (Exception ex)
            {
                AddLog(1, $"读取任务发生错误: {ex.Message}");
            }
        }

        private void UiDataGridView1_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < _variableList.Count)
            {
                var variable = _variableList[e.RowIndex];
                switch (uiDataGridView1.Columns[e.ColumnIndex].Name)
                {
                    case "VariableName":
                        e.Value = variable.VariableName;
                        break;
                    case "RevisionString":
                        e.Value = variable.RevisionString;
                        break;
                    case "AddressString":
                        e.Value = variable.AddressString;
                        break;
                    case "Size":
                        e.Value = variable.Size;
                        break;
                    case "Type":
                        e.Value = variable.Type;
                        break;
                    case "LowIndex":
                        e.Value = variable.LowIndex;
                        break;
                    case "HighIndex":
                        e.Value = variable.HighIndex;
                        break;
                    case "Value":
                        e.Value = variable.Value;
                        break;
                }
            }
        }

        private void UiDataGridView1_CellValuePushed(object sender, DataGridViewCellValueEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < _variableList.Count)
            {
                var variable = _variableList[e.RowIndex];
                switch (uiDataGridView1.Columns[e.ColumnIndex].Name)
                {
                    case "VariableName":
                        variable.VariableName = (string)e.Value;
                        break;
                    case "RevisionString":
                        // 不允许直接修改 RevisionString
                        break;
                    case "AddressString":
                        // 不允许直接修改 AddressString
                        break;
                    case "Size":
                        variable.Size = (int)e.Value;
                        break;
                    case "Type":
                        variable.Type = (string)e.Value;
                        break;
                    case "LowIndex":
                        variable.LowIndex = (int?)e.Value;
                        break;
                    case "HighIndex":
                        variable.HighIndex = (int?)e.Value;
                        break;
                    case "Value":
                        variable.Value = e.Value;
                        break;
                }
            }
        }


        private void UiDataGridView1_NewRowNeeded(object sender, DataGridViewRowEventArgs e)
        {
            // DataGridViewRowEventArgs does not have a Cancel property.  
            // To prevent new rows from being added, set the AllowUserToAddRows property of the DataGridView to false.  
            // 不允许添加新行
            uiDataGridView1.AllowUserToAddRows = false;
        }

        #endregion 通信调试
    }
}