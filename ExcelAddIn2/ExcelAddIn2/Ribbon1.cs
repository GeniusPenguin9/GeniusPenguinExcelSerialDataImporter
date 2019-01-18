using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace ExcelAddIn2
{
    public partial class Ribbon1
    {
        Excel.Application application;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            application = Globals.ThisAddIn.Application;
        }

        private void DB_Open_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "数据文件(.dat)|*.dat|所有文件(.)|*.*";
            openFileDialog.Multiselect = true;
            openFileDialog.Title = "选择数据文件";
            openFileDialog.ShowDialog();

            var line = 2;
            if (openFileDialog.FileName.Length > 0)
            {
                foreach (var file in openFileDialog.FileNames)
                {
                    for (int i = 0; i < 30; i++)
                    {
                        sheet.Cells[1, 3 + i * 7].Value = "Flow" + i;
                        sheet.Cells[1, 4 + i * 7].Value = "Speed" + i;
                        sheet.Cells[1, 5 + i * 7].Value = "Current" + i;
                        sheet.Cells[1, 6 + i * 7].Value = "IPressure" + i;
                        sheet.Cells[1, 7 + i * 7].Value = "OPressure" + i;
                        sheet.Cells[1, 8 + i * 7].Value = "SVoltage" + i;
                        sheet.Cells[1, 9 + i * 7].Value = "Alarm" + i;
                    }
                    var frames = ReadFramesFromFile(file);
                    uint StaticHead = 0xA050AA55;

                    sheet.Cells[1, 1].Value = "Start Time";
                    sheet.Cells[1, 2].Value = "End Time";
                   

                    int Old_Num = frames[0].Num - 1;
                    foreach (var frame in frames)
                    {
                        if (frame.Head == StaticHead)
                        {
                            ;
                            if (frame.Num == Old_Num + 1)
                            {
                                sheet.Cells[line, 1].Value = ConvertDateTime(frame.StartTime);
                                sheet.Cells[line, 2].Value = ConvertDateTime(frame.EndTime);
                                for (int i = 0; i < 30; i++)
                                {
                                    // 下记转换系数来源于AD转换系数，由硬件工程师提供
                                    sheet.Cells[line, 3 + i * 7].Value = frame.Data[i].Flow * 13.2 / 4096;
                                    sheet.Cells[line, 4 + i * 7].Value = frame.Data[i].Speed * 1.173;
                                    sheet.Cells[line, 5 + i * 7].Value = frame.Data[i].Current * 7.953 / 4096;
                                    sheet.Cells[line, 6 + i * 7].Value = (frame.Data[i].IPressure * 29.403 / 4096 - 14.7) * 51.7149;
                                    sheet.Cells[line, 7 + i * 7].Value = (frame.Data[i].OPressure * 29.403 / 4096 - 14.7) * 51.7149;
                                    sheet.Cells[line, 8 + i * 7].Value = frame.Data[i].SVoltage * 33.71 / 4096;
                                    sheet.Cells[line, 9 + i * 7].Value = frame.Data[i].Alarm;
                                }
                                line++;
                                Old_Num = frame.Num;
                            }
                            else {
                                MessageBox.Show("数据异常：存在漏帧");
                                return;
                            }
                        }
                        else
                        {
                            MessageBox.Show("数据异常：帧头错误");
                            return;
                        }
                    }
                }

                UInt32 ToggleEndian32(UInt32 number)
                {
                    return BitConverter.ToUInt32(BitConverter.GetBytes(number).Reverse().ToArray(), 0);
                }
                UInt16 ToggleEndian16(UInt16 number)
                {
                    return BitConverter.ToUInt16(BitConverter.GetBytes(number).Reverse().ToArray(), 0);
                }

                //将帧读出，计入列表
                List<Frame> ReadFramesFromFile(string file)
                {
                    var frameLength = 512;
                    var result = new List<Frame>();
                    var count = new FileInfo(file).Length / frameLength;
                    var reader = new BinaryReader(new FileStream(file, FileMode.Open));

                    for (int i = 0; i < count; i++)
                    {
                        var frame = new Frame();

                        // 16
                        frame.Head = reader.ReadUInt32();
                        frame.Num = ToggleEndian16(reader.ReadUInt16());
                        frame.StartTime = ToggleEndian32(reader.ReadUInt32());
                        frame.EndTime = ToggleEndian32(reader.ReadUInt32());
                        frame.Reserved1 = ToggleEndian16(reader.ReadUInt16());
                        frame.Data = new Data[30];

                        // 16*30=480
                        for (int j = 0; j < 30; j++)
                        {
                            frame.Data[j].Flow = ToggleEndian16(reader.ReadUInt16());
                            frame.Data[j].Speed = ToggleEndian16(reader.ReadUInt16());
                            frame.Data[j].Current = ToggleEndian16(reader.ReadUInt16());
                            frame.Data[j].IPressure = ToggleEndian16(reader.ReadUInt16());
                            frame.Data[j].OPressure = ToggleEndian16(reader.ReadUInt16());
                            frame.Data[j].SVoltage = ToggleEndian16(reader.ReadUInt16());
                            frame.Data[j].Reserved5 = ToggleEndian16(reader.ReadUInt16());
                            frame.Data[j].Alarm = ToggleEndian16(reader.ReadUInt16());
                        }

                        // 16
                        frame.Reserved2 = reader.ReadUInt32();
                        frame.Reserved3 = reader.ReadUInt32();
                        frame.Reserved4 = reader.ReadUInt32();
                        frame.Tail = reader.ReadUInt32();

                        result.Add(frame);
                    }

                    reader.Close();
                    return result;
                }

                //将其转换为符合规则的时间
                DateTime ConvertDateTime(UInt32 raw_time)
                {
                    return new DateTime(
                        (MaskByte(raw_time, 25, 31) + 1980),
                         MaskByte(raw_time, 21, 24),
                         MaskByte(raw_time, 16, 20),
                         MaskByte(raw_time, 11, 15),
                         MaskByte(raw_time, 5, 10),
                        (MaskByte(raw_time, 0, 4) * 2));
                }
                //提取指定位bit
                int MaskByte(UInt32 number, int start, int end)
                {
                    return (int)(number & (uint)(Math.Pow(2, end + 1) - Math.Pow(2, start))) >> start;
                }
            }
        }
        public struct Frame
        {
            public UInt32 Head;
            public UInt16 Num;
            public UInt32 StartTime;
            public UInt32 EndTime;
            public UInt16 Reserved1;
            public Data[] Data;
            public UInt32 Reserved2;
            public UInt32 Reserved3;
            public UInt32 Reserved4;
            public UInt32 Tail;
        }

        public struct Data
        {
            public UInt16 Flow;
            public UInt16 Speed;
            public UInt16 Current;
            public UInt16 IPressure;
            public UInt16 OPressure;
            public UInt16 SVoltage;
            public UInt16 Reserved5;
            public UInt16 Alarm;
        }

        private void DataCreate_Click(object sender, RibbonControlEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(XAxis.SelectedItem.Label) ||
                string.IsNullOrWhiteSpace(YAxis.SelectedItem.Label))
            { MessageBox.Show("请选择坐标轴"); return; }

            var sheet = (Excel.Worksheet)application.ActiveWorkbook.Worksheets.Add();
            sheet.Name = "图表生成";
            sheet.Range["A:A"].NumberFormatLocal = "yyyy/m/d h:mm";
            var irows = application.ActiveWorkbook.Worksheets["Sheet1"].UsedRange.Rows.Count;
            var RoundNum = Math.Floor((double)((irows - 1) / int.Parse( XAxis.SelectedItem.Tag)));
            for(var j = 0; j < RoundNum;j++)
            {
                sheet.Cells[j + 1,"A"]= 
                    application.ActiveWorkbook.Worksheets["Sheet1"].Cells[2 + j * int.Parse(XAxis.SelectedItem.Tag), 1];
                sheet.Cells[j + 1, "B"] =
                    application.ActiveWorkbook.Worksheets["Sheet1"].Cells[2 + j * int.Parse(XAxis.SelectedItem.Tag), YAxis.SelectedItem.Tag];
            }       
        }
    }
}
