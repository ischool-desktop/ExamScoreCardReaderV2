﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;

namespace ExamScoreCardReaderV2.Validation
{
    internal class ValidateTextFiles
    {
        //學號{7}，班級{3}，座號{2}，試別{2}，科目{5}，成績{6}
        //成績 = 000.00
        //        private const string format = @"^(\d{7})\d{3}\d{2}(\d{2}\d{5})([\d| ]\d{2}| [\d| ]\d)\.\d{2}$";
        // private const string format = @"^(\d{7})(\d{3})\d{2}(\d{2})(\d{5})(([\d| ]\d{2}| [\d| ]\d)\.\d{2})$";

        //學號
        private const string formatB = @"^([\d|\s]{";

        private const string formatE = @"})\s*([\d|\s]{3})\d{2}([\d|\s]{2})([\d|\s]{5})(([\d| ]\d{2}| [\d| ]\d)\.\d{2})$";

        // 2016/8/5 穎驊修正 上面那個formatE 的問題，注意其第二項([\d|\s]{2}) 沒有用()包起來，會使得後面判斷m.group 漏掉"座號"，造成錯誤
        //private const string formatE = @"})\s*([\d|\s]{3})([\d|\s]{2})([\d|\s]{2})([\d|\s]{5})(([\d| ]\d{2}| [\d| ]\d)\.\d{2})$";
        //2016/8/5 後來還是改回原本的，因為發現改了設定會造成其他驗證錯誤(EX 讀錯科目代碼) 要改還要研究其他的驗證功能，有點太大，恩正說先改回來

        private const string format2 = @"^%(.*):(\d+)%(.*)";

        //範例:   3350011 301 01 02 90001 78.45

        private Regex re;
        private Regex reline;

        public ValidateTextFiles(int StudentNumberLenght)
        {
            Global.TextFormat = @formatB + StudentNumberLenght + @formatE;

            re = new Regex(Global.TextFormat);
            reline = new Regex(format2);
        }

        internal ValidateTextResult CheckFormat(List<FileInfo> _files)
        {
            Dictionary<string, List<int>> uniqueCounter = new Dictionary<string, List<int>>();

            // 加這檢查主要處理匯入檔案內學號被不同班級座號重複使用
            Dictionary<string, List<string>> CheckStudNumberClassSeatNoDict = new Dictionary<string, List<string>>();


            //2016/8/5 穎驊新增 加這檢查主要處理匯入檔案內同一班級座號被不同學號重複使用
            Dictionary<string, List<string>> CheckClassSeat_StudNumberNoDict = new Dictionary<string, List<string>>();


            Dictionary<int, string> lineIndexes = new Dictionary<int, string>();
            List<int> errorFormatLineIndexes = new List<int>();
            List<int> duplicateLineIndexes = new List<int>();

            int index = 1;
            foreach (FileInfo file in _files)
            {
                int fileIndex = 1;
                lineIndexes.Add(index++, "[檔案名稱：" + file.Name + "]");
                StreamReader sr = new StreamReader(file.OpenRead());
                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    Match m = re.Match(line);
                    if (m.Success) //格式正確
                    {
                        //學號{7}，班級{3}，座號{2}，試別{2}，科目{5}
                        
                        //2016/8/5 穎驊註解，剛剛發現程式會遺漏掉座號，可以在這邊檢查每一個Value 是否是對的
                        string key = m.Groups[1].Value + m.Groups[2].Value + m.Groups[3].Value + m.Groups[4].Value + m.Groups[5].Value;

                        // 學號
                        string key1 = m.Groups[1].Value;
                        // 班座
                        string value1 = m.Groups[2].Value + m.Groups[3].Value;

                        if (!CheckStudNumberClassSeatNoDict.ContainsKey(key1))
                            CheckStudNumberClassSeatNoDict.Add(key1, new List<string>());

                        if (!CheckStudNumberClassSeatNoDict[key1].Contains(value1))
                            CheckStudNumberClassSeatNoDict[key1].Add(value1);

                        //2016/8/5 後來還是改回原本的，因為發現改了設定會造成其他驗證錯誤(EX 讀錯科目代碼) 要改還要研究其他的驗證功能，有點太大，恩正說先改回來
                        //if (!CheckClassSeat_StudNumberNoDict.ContainsKey(value1))
                        //    CheckClassSeat_StudNumberNoDict.Add(value1, new List<string>());

                        //if (!CheckClassSeat_StudNumberNoDict[value1].Contains(key1))
                        //    CheckClassSeat_StudNumberNoDict[value1].Add(key1);



                        if (!uniqueCounter.ContainsKey(key))
                        {
                            uniqueCounter.Add(key, new List<int>());
                            uniqueCounter[key].Add(index);
                        }
                        ////2016/8/5 穎驊註解， 重覆的話，幫你多加一筆index 資料，後面就會偵錯是否有>1筆表示重覆的錯誤，順帶一提 index 代表記事本的行數，index 1給 lineIndexes做表頭， index = 2 事實上等於記事本第1行
                        else {
                            uniqueCounter[key].Add(index);                        
                        }
                    }
                    else if (!string.IsNullOrEmpty(line.Trim())) //非空白行則錯誤
                        errorFormatLineIndexes.Add(index);
                    lineIndexes.Add(index++, "%" + file.Name + ":" + fileIndex + "%" + line);
                    fileIndex++;
                }
                sr.Close();
            }

            bool error = false;
            if (errorFormatLineIndexes.Count > 0)
            {
                error |= true;
                foreach (int key in errorFormatLineIndexes)
                {
                    if (errorFormatLineIndexes.Contains(key))
                        lineIndexes[key] += " (格式錯誤)";
                }
            }

            foreach (string key in uniqueCounter.Keys)
            {
                if (uniqueCounter[key].Count > 1) //學號+試別+科目 有重覆
                {
                    error |= true;
                    string f = " (學號+試別+科目重覆。{0})";
                    Dictionary<string, List<int>> group = new Dictionary<string, List<int>>();
                    foreach (int lineIndex in uniqueCounter[key])
                    {
                        Match m = reline.Match(lineIndexes[lineIndex]);
                        if (m.Success)
                        {
                            string fileName = m.Groups[1].Value;
                            string fileIndex = m.Groups[2].Value;
                            if (!group.ContainsKey(fileName))
                                group.Add(fileName, new List<int>());
                            group[fileName].Add(int.Parse(fileIndex));
                        }
                        if (!duplicateLineIndexes.Contains(lineIndex))
                            duplicateLineIndexes.Add(lineIndex);
                    }
                    string dupLines = string.Empty;

                    #region 先把重覆的行數隱藏起來，不顯示。
                    //foreach (string fileName in group.Keys)
                    //{
                    //    dupLines += "; " + fileName + ":";
                    //    foreach (int fileIndex in group[fileName])
                    //        dupLines += fileIndex + ",";
                    //    if (dupLines.EndsWith(",")) dupLines = dupLines.Substring(0, dupLines.Length - 1);
                    //}
                    //if (dupLines.StartsWith(";")) dupLines = dupLines.Substring(1, dupLines.Length - 1);
                    #endregion

                    foreach (int lineIndex in uniqueCounter[key])
                    {
                        if (lineIndexes.ContainsKey(lineIndex))
                            lineIndexes[lineIndex] += string.Format(f, dupLines);
                    }
                }
            }

            // 學號被不同人重複使用
            int errorIdx = 2;
            foreach (KeyValuePair<string, List<string>> data in CheckStudNumberClassSeatNoDict)
            {
                if (data.Value.Count > 1)
                {
                    error = true;
                    errorFormatLineIndexes.Add(errorIdx);
                    string msg = " (學號：" + data.Key + ", 有2位學生使用，請檢查!)";

                    if (lineIndexes.ContainsKey(errorIdx))
                        lineIndexes[errorIdx] += msg;
                    else
                        lineIndexes.Add(errorIdx, msg);

                }
                errorIdx++;
            }

            //2016/8/5 後來還是改回原本的，因為發現改了設定會造成其他驗證錯誤(EX 讀錯科目代碼) 要改還要研究其他的驗證功能，有點太大，恩正說先改回來
            ////座號 被不同人使用
            //int errorIdx_for_CheckClassSeat = 2;
            //foreach (KeyValuePair<string, List<string>> data in CheckClassSeat_StudNumberNoDict)
            //{
            //    if (data.Value.Count > 1)
            //    {
            //        error = true;
            //        errorFormatLineIndexes.Add(errorIdx_for_CheckClassSeat);
            //        string msg = " (班級座號：" + data.Key + ", 有2位學生使用，請檢查!)";

            //        if (lineIndexes.ContainsKey(errorIdx))
            //            lineIndexes[errorIdx_for_CheckClassSeat] += msg;
            //        else
            //            lineIndexes.Add(errorIdx_for_CheckClassSeat, msg);

            //    }
            //    errorIdx_for_CheckClassSeat++;
            //}


            ValidateTextResult result = new ValidateTextResult();

            if (error)
            {
                result.Error = true;
                result.LineIndexes = lineIndexes;
                result.ErrorFormatLineIndexes = errorFormatLineIndexes;
                result.DuplicateLineIndexes = duplicateLineIndexes;
            }

            return result;
        }
    }

    internal class ValidateTextResult
    {
        public bool Error { get; set; }
        public Dictionary<int, string> LineIndexes { get; set; }
        public List<int> ErrorFormatLineIndexes { get; set; }
        public List<int> DuplicateLineIndexes { get; set; }

        public ValidateTextResult()
        {
            Error = false;
        }
    }
}
