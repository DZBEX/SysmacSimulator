using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace SysmacSimulator
{
    public class SysmacSimulatorHelper
    {
        #region DLL导入

        [DllImport("C:\\Program Files\\OMRON\\Sysmac Studio\\MATLAB\\Win64\\NexSocket.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern void NexSock_initialize();

        [DllImport("C:\\Program Files\\OMRON\\Sysmac Studio\\MATLAB\\Win64\\NexSocket.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int NexSockClient_connect(ref short handle, byte[] ipAddress, short port);

        [DllImport("C:\\Program Files\\OMRON\\Sysmac Studio\\MATLAB\\Win64\\NexSocket.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int NexSock_close(short handle);

        [DllImport("C:\\Program Files\\OMRON\\Sysmac Studio\\MATLAB\\Win64\\NexSocket.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern void NexSock_terminate();

        [DllImport("C:\\Program Files\\OMRON\\Sysmac Studio\\MATLAB\\Win64\\NexSocket.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int NexSock_send(short handle, byte[] command, int length);

        [DllImport("C:\\Program Files\\OMRON\\Sysmac Studio\\MATLAB\\Win64\\NexSocket.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int NexSock_receive(short handle, byte[] buffer, int bufferSize);

        #endregion DLL导入

        /// 属性
        private string ipAddress = "127.0.0.1";

        private short port = 7000;
        private short handle;

        //变量字典
        public Dictionary<string, SimulatorVariable> variableDictionary = new Dictionary<string, SimulatorVariable>();

        /// <summary>
        /// 连接到模拟器
        /// </summary>
        public Tuple<bool, string> Connect()
        {
            try
            {
                NexSock_initialize();
                NexSockClient_connect(ref handle, Encoding.UTF8.GetBytes(ipAddress), port);
                return Tuple.Create(true, "连接成功!");
            }
            catch (Exception ex)
            {
                // 记录异常或处理异常
                string errorMsg = $" 执行  public bool Connect() 方法发生异常：{ex.Message}";
                return Tuple.Create(false, errorMsg);
            }
        }

        /// <summary>
        /// 断开连接
        /// </summary>
        public Tuple<bool, string> Disconnect()
        {
            try
            {
                NexSock_close(handle);
                NexSock_terminate();
                return Tuple.Create(true, "断开连接成功!");
            }
            catch (Exception ex)
            {
                // 记录异常或处理异常
                string errorMsg = $" 执行  public bool Disconnect() 方法发生异常：{ex.Message}";
                return Tuple.Create(false, errorMsg);
            }
        }

        /// <summary>
        /// 发送命令并接收响应
        /// </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        public Tuple<List<byte[]>, string> SendCommand(string command)
        {
            // 初始化响应列表和错误变量
            List<byte[]> response = new List<byte[]>();
            string error = null;

            // 创建一个512字节的缓冲区
            byte[] buffer = new byte[512];
            try
            {
                // 调用NexSock_send函数发送命令
                int sendResult = NexSock_send(handle, Encoding.UTF8.GetBytes(command), command.Length);
                if (sendResult < 0)
                {
                    error = "Failed to send command";
                    return Tuple.Create(response, error);
                }

                // 循环接收响应，直到响应长度为0
                while (true)
                {
                    int responseLength = NexSock_receive(handle, buffer, buffer.Length);
                    if (responseLength == 0)
                    {
                        // 如果响应长度为0，表示接收完毕，跳出循环
                        break;
                    }
                    else if (responseLength < 0)
                    {
                        // 如果响应长度小于0，表示接收出错，记录错误信息
                        error = Encoding.UTF8.GetString(buffer).TrimEnd('\0');
                        break;
                    }
                    else
                    {
                        // 如果响应长度大于0，表示接收到有效响应，存储响应信息
                        //response.Add(Encoding.UTF8.GetString(buffer, 0, responseLength).TrimEnd('\0'));
                        // 创建目标数组
                        byte[] bufferResponse = new byte[responseLength];
                        // 使用 Array.Copy 截取子数组
                        Array.Copy(buffer, 0, bufferResponse, 0, responseLength);
                        response.Add(bufferResponse);
                    }
                }
                // 返回响应列表和错误变量
                return Tuple.Create(response, error);
            }
            catch (Exception ex)
            {
                // 记录异常或处理异常
                string errorMsg = $" 执行  public Tuple<List<byte[]>, string> SendCommand(string command) 方法发生异常：{ex.Message}";
                // 返回响应列表和错误变量
                return Tuple.Create(response, error);
            }
        }

        #region 读取变量

        /// <summary>
        /// 读取变量
        /// </summary>
        /// <param name="variableName"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public Tuple<bool, string, object> ReadVariable(string variableName)
        {
            try
            {
                // 检查变量是否已在缓存中，若不在则通过变量名获取变量信息并添加到缓存
                if (!variableDictionary.ContainsKey(variableName))
                {
                    variableDictionary[variableName] = GetVariableInfo(variableName);
                }

                // 从缓存中获取变量信息
                SimulatorVariable variableInfo = variableDictionary[variableName];

                // 构造异步读取内存文本的命令，并发送命令获取响应
                string command = $"AsyncReadMemText {Encoding.UTF8.GetString(variableInfo.Revision).TrimEnd()} 1 {Encoding.UTF8.GetString(variableInfo.Address).TrimEnd()},2";
                Console.WriteLine("时间：" + DateTime.Now.ToString("HH:mm:ss:ffff") + "读取指令:" + variableName + "[ " + command + " ]");
                var responseTuple = SendCommand(command);
                byte[] responseData = responseTuple.Item1[0];
                Console.WriteLine("时间：" + DateTime.Now.ToString("HH:mm:ss:ffff") + "返回响应:" + variableName + "[ " + BitConverter.ToString(responseData) + " ]");
                var error = responseTuple.Item2;
                if (error != null)
                {
                    return Tuple.Create<bool, string, object>(false, variableName + ":" + $"Error reading variable: {error}", null);
                }
                // 根据变量类型解析响应数据，并返回解析后的值
                object result = UnpackBytes(responseData, variableInfo.Type);
                return Tuple.Create(true, "读取成功!" + "   " + variableName + "   -   " + result, result);
            }
            catch (Exception ex)
            {
                // 记录异常或处理异常
                string errorMsg = $" 执行  public Tuple<bool, string> WriteVariable(string variableName, object value) 方法发生异常：{ex.Message}";
                return Tuple.Create<bool, string, object>(false, variableName + ":" + $"Error writing variable: {errorMsg}", null);
            }
        }

        #endregion 读取变量

        #region 写入变量

        /// <summary>
        /// 写入变量
        /// </summary>
        /// <param name="variableName"></param>
        /// <param name="value"></param>
        /// <exception cref="Exception"></exception>
        public Tuple<bool, string> WriteVariable(string variableName, object value)
        {
            try
            {
                // 检查变量是否已在缓存中，若不在则获取变量信息并加入缓存
                if (!variableDictionary.ContainsKey(variableName))
                {
                    variableDictionary[variableName] = GetVariableInfo(variableName);
                }
                SimulatorVariable variableInfo = variableDictionary[variableName];
                // 将要写入的值打包为字节流
                byte[] sendBytes = PackBytes(value, variableInfo.Type, variableInfo.Size);
                // 构造写入命令
                string command = $"AsyncWriteMemText {Encoding.UTF8.GetString(variableInfo.Revision).TrimEnd()} 1 {Encoding.UTF8.GetString(variableInfo.Address).TrimEnd()},2,{BitConverter.ToString(sendBytes).Replace("-", "")}";
                Console.WriteLine("时间：" + DateTime.Now.ToString("HH:mm:ss:ffff") + "写入指令:" + variableName + "[ " + command + " ]");
                // 发送写入命令并返回响应
                var responseTuple = SendCommand(command);
                var response = responseTuple.Item1;
                var error = responseTuple.Item2;
                if (error != null)
                {
                    return Tuple.Create(false, "写入失败!" + "   " + variableName + "   -   " + value + $"Error writing variable: {error}");
                }
                return Tuple.Create(true, "写入成功!" + "   " + variableName + "   -   " + value);
            }
            catch (Exception ex)
            {
                // 记录异常或处理异常
                string errorMsg = $" 执行  public Tuple<bool, string> WriteVariable(string variableName, object value) 方法发生异常：{ex.Message}";
                return Tuple.Create(false, variableName + ":" + value + "写入失败!" + $"Error writing variable: {errorMsg}");
            }
        }

        #endregion 写入变量

        /// <summary>
        /// 将数据打包为字节流
        /// </summary>
        /// <param name="data"></param>
        /// <param name="dataType"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        /// <exception cref="NotSupportedException"></exception>
        private byte[] PackBytes(object data, string dataType, int size)
        {
            byte[] packedBytes = new byte[size];
            // 根据不同的数据类型进行打包处理
            if (dataType == "BOOL")
            {
                bool value = bool.Parse(((string)data).ToLower());
                if (value)
                {
                    packedBytes[0] = 255;
                }
                else
                {
                    packedBytes[0] = 0;
                }
            }
            else if (dataType == "SINT")
            {
                if (short.TryParse((string)data, out short value))
                {
                    byte[] MypackedBytes = BitConverter.GetBytes(value);
                    packedBytes[0] = MypackedBytes[0];
                }
            }
            else if (dataType == "INT")
            {
                if (short.TryParse((string)data, out short value))
                {
                    packedBytes = BitConverter.GetBytes(value);
                }
            }
            else if (dataType == "DINT")
            {
                if (int.TryParse((string)data, out int value))
                {
                    packedBytes = BitConverter.GetBytes(value);
                }
            }
            else if (dataType == "LINT")
            {
                if (long.TryParse((string)data, out long value))
                {
                    packedBytes = BitConverter.GetBytes(value);
                }
            }
            else if (dataType == "USINT")
            {
                if (ushort.TryParse((string)data, out ushort value))
                {
                    byte[] MypackedBytes = BitConverter.GetBytes(value);
                    packedBytes[0] = MypackedBytes[0];
                }
            }
            else if (dataType == "UINT")
            {
                if (ushort.TryParse((string)data, out ushort value))
                {
                    packedBytes = BitConverter.GetBytes(value);
                }
            }
            else if (dataType == "UDINT")
            {
                if (uint.TryParse((string)data, out uint value))
                {
                    packedBytes = BitConverter.GetBytes(value);
                }
            }
            else if (dataType == "ULINT")
            {
                if (ulong.TryParse((string)data, out ulong value))
                {
                    packedBytes = BitConverter.GetBytes(value);
                }
            }
            else if (dataType == "REAL")
            {
                if (float.TryParse((string)data, out float value))
                {
                    packedBytes = BitConverter.GetBytes(value);
                }
            }
            else if (dataType == "LREAL")
            {
                if (double.TryParse((string)data, out double value))
                {
                    packedBytes = BitConverter.GetBytes(value);
                }
            }
            else if (dataType.StartsWith("STRING"))
            {
                byte[] stringBytes = Encoding.UTF8.GetBytes(data.ToString());
                Array.Copy(stringBytes, packedBytes, Math.Min(stringBytes.Length, size));
            }
            else
            {
                // 对于未知的数据类型，不进行处理
                throw new NotSupportedException($"Unsupported data type: {dataType}");
            }
            return packedBytes;
        }

        /// <summary>
        /// 解包字节数组
        /// </summary>
        /// <param name="data"></param>
        /// <param name="dataType"></param>
        /// <returns></returns>
        /// <exception cref="NotSupportedException"></exception>
        private object UnpackBytes(byte[] data, string dataType)
        {
            object result = null;
            // 根据不同的数据类型进行解包处理
            if (dataType == "BOOL")
            {
                // 对BOOL类型数据进行解包
                result = BitConverter.ToBoolean(data, 0);
            }
            else if (dataType == "SINT")
            {
                byte[] extendedData = new byte[2]; // 创建一个4字节的数组
                extendedData[0] = data[0]; // 将单字节数据复制到第一个位置
                result = BitConverter.ToInt16(extendedData, 0);
            }
            else if (dataType == "INT")
            {
                // 对INT类型数据进行解包
                result = BitConverter.ToInt16(data, 0);
            }
            else if (dataType == "DINT")
            {
                // 对DINT类型数据进行解包
                result = BitConverter.ToInt32(data, 0);
            }
            else if (dataType == "LINT")
            {
                // 对LINT类型数据进行解包
                result = BitConverter.ToInt64(data, 0);
            }
            else if (dataType == "USINT")
            {
                // 对USINT类型数据进行解包
                byte[] extendedData = new byte[2]; // 创建一个4字节的数组
                extendedData[0] = data[0]; // 将单字节数据复制到第一个位置
                result = BitConverter.ToUInt16(extendedData, 0);
            }
            else if (dataType == "UINT")
            {
                // 对UINT类型数据进行解包
                result = BitConverter.ToUInt16(data, 0);
            }
            else if (dataType == "UDINT")
            {
                // 对UDINT类型数据进行解包
                result = BitConverter.ToUInt32(data, 0);
            }
            else if (dataType == "ULINT")
            {
                // 对ULINT类型数据进行解包
                result = BitConverter.ToUInt64(data, 0);
            }
            else if (dataType == "REAL")
            {
                // 对REAL类型数据进行解包
                result = BitConverter.ToSingle(data, 0);
            }
            else if (dataType == "LREAL")
            {
                // 对LREAL类型数据进行解包
                result = BitConverter.ToDouble(data, 0);
            }
            else if (dataType.StartsWith("STRING"))
            {
                // 对STRING类型数据进行解包
                result = Encoding.UTF8.GetString(data).TrimEnd('\0');
            }
            else
            {
                // 对于未知的数据类型，不进行处理
                throw new NotSupportedException($"Unsupported data type: {dataType}");
            }
            // 返回解包后的数据
            return result;
        }

        #region 全局变量文件解析

        /// <summary>
        /// 从文件中读取变量信息
        /// </summary>
        /// <param name="filename"></param>
        public void PopulateFromFile(string filename)
        {
            // 打开文件，准备逐行读取并处理数据
            using (StreamReader file = new StreamReader(filename))
            {
                // 跳过文件的第一行，可能是标题或不需要的信息
                file.ReadLine();
                // 遍历文件的每一行，提取变量信息
                string line;
                while ((line = file.ReadLine()) != null)
                {
                    // 分割行数据，去除末尾的换行符
                    string[] tokens = line.Split(new char[] { '\t' }, StringSplitOptions.RemoveEmptyEntries); //['Value', 'REAL', 'TRUE', 'RW'] //['Value', 'REAL', 'TRUE', 'RW']
                    // 确保当前行包含变量信息
                    if (tokens.Length > 1)
                    {
                        // 检查是否存在索引，表示可能是数组或序列
                        if (Regex.IsMatch(tokens[1], @"\[.*?\]"))
                        {
                            // 提取索引信息和数据类型
                            MatchCollection indexes = Regex.Matches(tokens[1], @"\[.*?\]");
                            string dataType = Regex.Replace(tokens[1], @"\[.*?\]", "");
                            // 处理存在索引的情况
                            if (indexes.Count > 0)
                            {
                                string[] indexRange = Regex.Replace(indexes[0].Value, @"[\[\]{}]", "").Split(new string[] { ".." }, StringSplitOptions.RemoveEmptyEntries);
                                int index = int.Parse(indexRange[0]);
                                int end = int.Parse(indexRange[1]);
                                // 根据索引范围创建变量并添加到字典
                                while (index <= end)
                                {
                                    string variableName = tokens[0] + $"[{index}]";
                                    SimulatorVariable simVariable = GetVariableInfo(variableName);
                                    simVariable.Type = dataType;
                                    variableDictionary[variableName] = simVariable;
                                    index++;
                                }
                            }
                        }
                        else
                        {
                            // 处理不存在索引的情况，直接创建变量并添加到字典
                            SimulatorVariable simVariable = GetVariableInfo(tokens[0]);
                            simVariable.Type = tokens[1];
                            variableDictionary[tokens[0]] = simVariable;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 获取变量信息
        /// </summary>
        /// <param name="variableName"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public SimulatorVariable GetVariableInfo(string variableName)
        {
            // 发送命令以获取变量地址和大小信息
            var command = $"GetVarAddrText 1 VAR://{variableName}";
            var responseTuple = SendCommand(command);
            var response = responseTuple.Item1;
            var error = responseTuple.Item2;

            if (error != null)
            {
                throw new Exception($"Error getting variable info: {error}");
            }

            // 解码并提取变量的修订号
            byte[] revision = response[0];

            // 解码并提取变量的地址信息，去除末尾的换行符
            byte[] address = response[2];

            // 从地址信息中计算变量的大小（以字节为单位）
            string receivedString = Encoding.UTF8.GetString(response[2]).TrimEnd();
            string[] parts = receivedString.Split(',');
            int isize = int.Parse(parts[parts.Length - 1]) / 8;
            if (isize == 0)
            {
                isize = 1;
            }
            int size = isize;
            string value = null;
            // 创建并返回变量信息对象
            var variableInfo = new SimulatorVariable(variableName, revision, address, size, value);
            return variableInfo;
        }

        #endregion 全局变量文件解析
    }
}