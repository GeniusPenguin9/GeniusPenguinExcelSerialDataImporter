using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
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
            Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "数据文件(.db)|*.db|所有文件(.)|*.*";
            openFileDialog.Multiselect = true;
            openFileDialog.Title = "选择数据文件";
            openFileDialog.ShowDialog();
            if (openFileDialog.FileName.Length > 0)
            {
                foreach (var file in openFileDialog.FileNames)
                {
                    var frames = ReadFramesFromFile(file);
                    var line = 2;

                    foreach (var frame in frames)
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
        }

            public List<Frame> ReadFramesFromFile(string file)
                {
                    var frameLength = 188;
                    var result = new List<Frame>();
                    var count = new FileInfo(file).Length / frameLength;
                    var reader = new BinaryReader(new FileStream(file, FileMode.Open));

                    for (int i = 0; i < count; i++)
                    {
                        var frame = new Frame();

                        frame.Head = reader.ReadInt16();
                        frame.StartTime = reader.ReadInt16();
                        frame.EndTime = reader.ReadInt16();
                        frame.Data = new Data[30];
                        for (int j = 0; j < 30; j++)
                        {
                            frame.Data[j].Flow = reader.ReadDouble();
                            frame.Data[j].Speed = reader.ReadDouble();
                            frame.Data[j].Voltage = reader.ReadDouble();
                        }
                        frame.Tail = reader.ReadInt16();

                        result.Add(frame);
                    }

                    reader.Close();
                    return result;
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

//namespace ExcelAddIn1
//{
//    public struct Frame
//    {
//        public Int16 Head;
//        public Int16 StartTime;
//        public Int16 EndTime;
//        public Data[] Data;
//        public Int16 Tail;
//    }

//    public struct Data
//    {
//        public double Flow;
//        public double Speed;
//        public double Voltage;
//    }

//    public partial class Ribbon1
//    {
//        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
//        {
//        }

//        private void button1_Click(object sender, RibbonControlEventArgs e)
//        {
//            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;

//            var fd = new OpenFileDialog();
//            fd.Filter = "dat文件|*.dat";
//            fd.Multiselect = true;
//            fd.ShowDialog();
//            if (fd.FileNames.Length > 0)
//            {
//                foreach (var file in fd.FileNames)
//                {
//                    var frames = ReadFramesFromFile(file);
//                    var line = 2;

//                    foreach (var frame in frames)
//                    {

//                        sheet.Cells[line, 1].Value = frame.StartTime;
//                        sheet.Cells[line, 2].Value = frame.EndTime;
//                        for (int i = 0; i < 30; i++)
//                        {
//                            sheet.Cells[line, 3 + i * 3].Value = frame.Data[i].Flow;
//                            sheet.Cells[line, 4 + i * 3].Value = frame.Data[i].Speed;
//                            sheet.Cells[line, 5 + i * 3].Value = frame.Data[i].Voltage;
//                        }

//                        line++;
//                    }

//                }
//            }
//        }

//        private List<Frame> ReadFramesFromFile(string file)
//        {
//            var frameLength = 188;
//            var result = new List<Frame>();
//            var count = new FileInfo(file).Length / frameLength;
//            var reader = new BinaryReader(new FileStream(file, FileMode.Open));

//            for (int i = 0; i < count; i++)
//            {
//                var frame = new Frame();

//                frame.Head = reader.ReadInt16();
//                frame.StartTime = reader.ReadInt16();
//                frame.EndTime = reader.ReadInt16();
//                frame.Data = new Data[30];
//                for (int j = 0; j < 30; j++)
//                {
//                    frame.Data[j].Flow = reader.ReadDouble();
//                    frame.Data[j].Speed = reader.ReadDouble();
//                    frame.Data[j].Voltage = reader.ReadDouble();
//                }
//                frame.Tail = reader.ReadInt16();

//                result.Add(frame);
//            }

//            reader.Close();
//            return result;
//        }
//    }
//}