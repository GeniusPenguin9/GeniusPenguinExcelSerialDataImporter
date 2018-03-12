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
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void DB_Open_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "数据文件(.dat)|*.dat|所有文件(.)|*.*";
            openFileDialog.Multiselect = true;
            openFileDialog.Title = "选择数据文件";
            openFileDialog.ShowDialog();
            if (openFileDialog.FileName.Length > 0)
            {
                foreach (var file in openFileDialog.FileNames)
                {
                    var frames = ReadFramesFromFile(file);
                    var line = 2;
                    uint StaticHead = 0xA050AA55;

                    sheet.Cells[1, 1].Value = "Start Time";
                    sheet.Cells[1, 2].Value = "End Time";
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

                    foreach (var frame in frames)
                    {
                        if (frame.Head == StaticHead)
                        {
                            sheet.Cells[line, 1].Value = ConvertDateTime(frame.StartTime);
                            sheet.Cells[line, 2].Value = ConvertDateTime(frame.EndTime);
                            for (int i = 0; i < 30; i++)
                            {
                                sheet.Cells[line, 3 + i * 7].Value = frame.Data[i].Flow;
                                sheet.Cells[line, 4 + i * 7].Value = frame.Data[i].Speed;
                                sheet.Cells[line, 5 + i * 7].Value = frame.Data[i].Current;
                                sheet.Cells[line, 6 + i * 7].Value = frame.Data[i].IPressure;
                                sheet.Cells[line, 7 + i * 7].Value = frame.Data[i].OPressure;
                                sheet.Cells[line, 8 + i * 7].Value = frame.Data[i].SVoltage;
                                sheet.Cells[line, 9 + i * 7].Value = frame.Data[i].Alarm;
                            }

                            line++;
                        }
                        else
                        {
                            MessageBox.Show("数据错误");
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
                        frame.Reserved1 = reader.ReadUInt16();
                        frame.Data = new Data[30];

                        // 16*30=480
                        for (int j = 0; j < 30; j++)
                        {
                            frame.Data[j].Flow = reader.ReadInt16();
                            frame.Data[j].Speed = reader.ReadInt16();
                            frame.Data[j].Current = reader.ReadInt16();
                            frame.Data[j].IPressure = reader.ReadInt16();
                            frame.Data[j].OPressure = reader.ReadInt16();
                            frame.Data[j].SVoltage = reader.ReadInt16();
                            frame.Data[j].Reserved5 = reader.ReadInt16();
                            frame.Data[j].Alarm = reader.ReadInt16();
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

                DateTime ConvertDateTime(UInt32 raw_time)
                {
                    // TODO: delete test
                    var test = new List<int>
                    {
                        MaskByte(raw_time, 25, 31),
                        MaskByte(raw_time, 24, 21),
                        MaskByte(raw_time, 20, 16),
                        MaskByte(raw_time, 15, 11),
                        MaskByte(raw_time, 10, 5),
                        MaskByte(raw_time, 4, 0)
                    };
                    return new DateTime(
                        MaskByte(raw_time, 25, 31),
                        MaskByte(raw_time, 24, 21),
                        MaskByte(raw_time, 20, 16),
                        MaskByte(raw_time, 15, 11),
                        MaskByte(raw_time, 10, 5),
                        MaskByte(raw_time, 4, 0));
                }

                int MaskByte(UInt32 number, int start, int end)
                {
                    return (int)(number & (int)(Math.Pow( 2, end) - Math.Pow(2, start))) >> start;
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
            public Int16 Flow;
            public Int16 Speed;
            public Int16 Current;
            public Int16 IPressure;
            public Int16 OPressure;
            public Int16 SVoltage;
            public Int16 Reserved5;
            public Int16 Alarm;
        }
    }
}
