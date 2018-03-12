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
                    uint StaticHead = 0xAA55A050;

                    foreach (var frame in frames)
                    {
                        var Head2 = frame.Head;
                        if (frame.Head == StaticHead)
                        {
                            sheet.Cells[line, 1].Value = frame.StartTime;
                            sheet.Cells[line, 2].Value = frame.EndTime;
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
                        { MessageBox.Show("数据错误");
                        }
                    }
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

                        frame.Head = reader.ReadInt32();
                        frame.Num = reader.ReadInt16();
                        frame.StartTime = reader.ReadInt32();
                        frame.EndTime = reader.ReadInt32();
                        frame.Reserved1 = reader.ReadInt16();
                        frame.Data = new Data[30];
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
                        frame.Reserved2 = reader.ReadInt16();
                        frame.Reserved3 = reader.ReadInt16();
                        frame.Reserved4 = reader.ReadInt16();
                        frame.Tail = reader.ReadInt16();

                        result.Add(frame);
                    }

                    reader.Close();
                    return result;
                }
            }
            
            }
        public struct Frame
        {
            public int Head;
            public short Num;
            public int StartTime;
            public int EndTime;
            public short Reserved1;
            public Data[] Data;
            public int Reserved2;
            public int Reserved3;
            public int Reserved4;
            public int Tail;
        }

        public struct Data
        {
            public short Flow;
            public short Speed;
            public short Current;
            public short IPressure;
            public short OPressure;
            public short SVoltage;
            public short Reserved5;
            public short Alarm;
        }
    }
}
