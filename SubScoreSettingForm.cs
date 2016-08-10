using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.Xml.Serialization;

namespace ExamScoreCardReaderV2
{
    public partial class SubScoreSettingForm : BaseForm
    {

        // 取得CourseList
        List<SP_CourseRecord> _CourseList = SP_course.SelectByIDs(K12.Presentation.NLDPanels.Course.SelectedSource);

        // 紀錄Exam_ID 之 Dictionary
        Dictionary<int, String> _DicExamID = new Dictionary<int, string>();

        public SubScoreSettingForm()
        {
            InitializeComponent();

            FISCA.Data.QueryHelper qh = new FISCA.Data.QueryHelper();

            // 因為K12 無法取得Extension欄位的資料(其為internal)，所以另外下SQL 抓Table 使用
            string strSQL_course = "SELECT * FROM course WHERE id IN (" + string.Join(",", K12.Presentation.NLDPanels.Course.SelectedSource.ToArray()) + ")";

            System.Data.DataTable dt_course = qh.Select(strSQL_course);

            //抓取評量設定
            string strSQL_exam = "SELECT * FROM exam";

            System.Data.DataTable dt_exam = qh.Select(strSQL_exam);

            //為 DataGridView 新增評量Col
            if (dt_exam.Rows.Count != 0)
            {
                for (int ii = 0; ii < dt_exam.Rows.Count; ii++)
                {
                    DataGridViewCheckBoxColumn col = new DataGridViewCheckBoxColumn();
                    col.CellTemplate = new DataGridViewCheckBoxCell();

                    col.Name = dt_exam.Rows[ii]["exam_name"].ToString();
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
                    col.Visible = true;
                    col.Width = 140;

                    // 將Exam_ID 加入，Key 使用ii +3 是因為方便後續取得Exam_id使用，可以直接填入Col = 3 、4、5 來查詢(UI 上 子成績項目是從第三欄開始)
                    _DicExamID.Add(ii + 3, dt_exam.Rows[ii]["id"].ToString());

                    dataGridViewX1.Columns.Add(col);
                }
            }

            // 將Course 內的課程資訊加到dataGridViewX1
            if (_CourseList.Count != 0)
            {
                foreach (var course in _CourseList)
                {
                    DataGridViewRow row = new DataGridViewRow();

                    // 其實這邊可以直接指定後面每一項ChkBox 的值 像是row.CreateCells(dataGridViewX1, course.SchoolYear, course.Semester, course.Name,true,true,false ....);
                    // 但想了想，這樣做就無法動態產生，所以會有後續的Code
                    row.CreateCells(dataGridViewX1, course.SchoolYear, course.Semester, course.Name);

                    for (int i = 0; i < dt_course.Rows.Count; i++)
                    {
                        if (course.ID == dt_course.Rows[i]["id"].ToString())
                        {
                            XmlDocument doc1 = new XmlDocument();
                            string strXml = "" + dt_course.Rows[i]["Extensions"];

                            if (strXml != "")
                            {
                                doc1.LoadXml(strXml);

                                XmlElement ele = doc1.DocumentElement;

                                XmlElement eleGradeItemExtension = ele.SelectSingleNode("Extension[@Name='GradeItemExtension']") as XmlElement;

                                //需要檢驗eleGradeItemExtension 是不是null， 因為Extensions 內可能會沒有Extension[@Name='GradeItemExtension' 此類作為子成績使用的項目Element ，但卻有其他功能的東西，或是亂塞的東西
                                if (eleGradeItemExtension != null)
                                {
                                    for (int exam_col = 0; exam_col < dt_exam.Rows.Count; exam_col++)
                                    {
                                        XmlElement eleGradeItem = eleGradeItemExtension.SelectSingleNode("GradeItemExtension[@ExamID=" + "'" + _DicExamID[exam_col + 3] + "'" + "]") as XmlElement;

                                        // 去選取有該Exam ID 的XmlElement ，找得到代表有該項子成績，Chk = true，反之Chk = false
                                        if (eleGradeItem != null)
                                        {
                                            row.Cells[exam_col + 3].Value = true;
                                        }
                                        else
                                        {
                                            row.Cells[exam_col + 3].Value = false;
                                        }
                                    }
                                }

                                // eleGradeItemExtension 為null 的話，代表完全沒有子成績項目，所以所有子成績要設為False
                                else
                                {
                                    for (int exam_col = 0; exam_col < dt_exam.Rows.Count; exam_col++)
                                    {
                                        row.Cells[exam_col + 3].Value = false;
                                    }
                                }
                                dataGridViewX1.Rows.Add(row);
                            }
                            // 假如抓下的 String (strXml = "" + dt_course.Rows[i]["Extensions"]) 為null 的話，代表完全沒有子成績、Extensions 欄位裏頭甚麼都沒有，所以所有子成績要設為False
                            else
                            {
                                for (int exam_col = 0; exam_col < dt_exam.Rows.Count; exam_col++)
                                {
                                    row.Cells[exam_col + 3].Value = false;
                                }
                                dataGridViewX1.Rows.Add(row);
                            }
                        }
                    }
                }
            }
        }


        //儲存設定
        private void buttonX1_Click(object sender, EventArgs e)
        {
            FISCA.Data.QueryHelper qh = new FISCA.Data.QueryHelper();

            UpdateHelper uh = new UpdateHelper();

            foreach (SP_CourseRecord SPCR in _CourseList)
            {
                foreach (DataGridViewRow dgvr in dataGridViewX1.Rows)
                {
                    if (dgvr.Cells[0].Value != null && dgvr.Cells[1].Value != null && dgvr.Cells[2].Value != null)
                    {
                        if (SPCR.SchoolYear.ToString() == dgvr.Cells[0].Value.ToString() && SPCR.Semester.Value.ToString() == dgvr.Cells[1].Value.ToString() && SPCR.Name == dgvr.Cells[2].Value.ToString())
                        {
                            String body = "";

                            string strSQL_course = "SELECT * FROM course WHERE id =" + SPCR.ID;

                            System.Data.DataTable dt_course = qh.Select(strSQL_course);

                            XmlDocument doc1 = new XmlDocument();

                            string strXml = "" + dt_course.Rows[0]["Extensions"];

                            // 如果讀下來的Course Row 其Extension欄位 不為""，LoadXml， 否則就幫他創
                            if (strXml != "")
                            {
                                doc1.LoadXml(strXml);
                            }

                            // 恩正說Extensions 裡頭的格式如下(把註解符號弄掉就是原本的樣子)，因應這樣的格式所以會有後面XmlDocument、XmlElement 的處理
                            //                            String body = @" '<Extensions>
                            //	                                        <Extension Name= ""GradeItemExtension"" >
                            //	                                         <GradeItemExtension Calc=""SUM"" ExamID=""1"">
                            //		                                        <SubExam ExtName=""CScore"" Permission=""Read"" SubName=""讀卡"" Type=""Number""/>
                            //		                                        <SubExam ExtName=""PScore"" Permission=""Editor"" SubName=""試卷"" Type=""Number""/>
                            //	                                        </GradeItemExtension>
                            //	                                         <GradeItemExtension Calc=""SUM"" ExamID=""3"">
                            //		                                        <SubExam ExtName=""CScore"" Permission=""Read"" SubName=""讀卡""Type=""Number""/>
                            //		                                        <SubExam ExtName=""PScore"" Permission=""Editor"" SubName=""試卷"" Type=""Number""/>
                            //	                                            </GradeItemExtension>
                            //	                                            </Extension>
                            //                                            </Extensions>'";

                            //指定 ele 為 Xmldocument doc1 的Element
                            XmlElement ele = doc1.DocumentElement;

                            //  假如指定後，ele 沒有東西，代表原本Xmldocument doc1的內容物也是沒有東西，就幫他加東西
                            if (ele == null)
                            {

                                XmlElement extensionsHead = doc1.CreateElement("Extensions");

                                XmlElement extensionsBody = doc1.CreateElement("Extension");

                                extensionsHead.AppendChild(extensionsBody);

                                extensionsBody.SetAttribute("Name", "GradeItemExtension");

                                for (int i = 3; i < dataGridViewX1.ColumnCount; i++)
                                {
                                    if (dgvr.Cells[i].Value != null)
                                        if (dgvr.Cells[i].Value.ToString() == "True")
                                        {
                                            XmlElement GradeItemExtension = doc1.CreateElement("GradeItemExtension");

                                            extensionsBody.AppendChild(GradeItemExtension);

                                            GradeItemExtension.SetAttribute("Calc", "SUM");

                                            GradeItemExtension.SetAttribute("ExamID", _DicExamID[i]);

                                            XmlElement SubExam1 = doc1.CreateElement("SubExam");

                                            XmlElement SubExam2 = doc1.CreateElement("SubExam");

                                            GradeItemExtension.AppendChild(SubExam1);

                                            GradeItemExtension.AppendChild(SubExam2);

                                            SubExam1.SetAttribute("ExtName", "CScore");

                                            SubExam1.SetAttribute("Permission", "Read");

                                            SubExam1.SetAttribute("SubName", "讀卡");

                                            SubExam1.SetAttribute("Type", "Number");

                                            SubExam2.SetAttribute("ExtName", "PScore");

                                            SubExam2.SetAttribute("Permission", "Editor");

                                            SubExam2.SetAttribute("SubName", "試卷");

                                            SubExam2.SetAttribute("Type", "Number");

                                        }
                                    //  這裡不做dgvr.Cells[i].Value.ToString() == "Flase" 的處理，是因為 假如原本Extensions欄位是空的話，只需要新增，不需要移除
                                }

                                // 因為ele 沒有東西，所以傳新建立XmlElement extensionsHead的的進去   記得要下SQL 內的String 要用 ' ' 包起來ㄚㄚ~
                                body = "'" + extensionsHead.OuterXml + "'";
                            }
                            // 假如原本Extensions欄位有東西(ele != null)，再另外處理
                            else
                            {
                                for (int i = 3; i < dataGridViewX1.ColumnCount; i++)
                                {
                                    if (dgvr.Cells[i].Value != null)
                                        if (dgvr.Cells[i].Value.ToString() == "True")
                                        {
                                            XmlElement eleGradeItemExtension = ele.SelectSingleNode("Extension[@Name='GradeItemExtension']") as XmlElement;

                                            // eleGradeItemExtension 有可能會=null， 代表原本Extensions欄位有其他奇怪的東西(EX: <Abbbbbs><Abbbbb>aaaa</Abbbbb></Abbbbbs>)
                                            if (eleGradeItemExtension != null)
                                            {
                                                XmlElement eleGradeItem = eleGradeItemExtension.SelectSingleNode("GradeItemExtension[@ExamID=" + "'" + _DicExamID[i] + "'" + "]") as XmlElement;

                                                // 在子成績CheckBox = true 的狀況下，去搜尋那一項的子成績已經加入，如果沒有 就幫他創。
                                                if (eleGradeItem == null)
                                                {
                                                    XmlElement GradeItemExtension = doc1.CreateElement("GradeItemExtension");

                                                    eleGradeItemExtension.AppendChild(GradeItemExtension);

                                                    GradeItemExtension.SetAttribute("Calc", "SUM");

                                                    GradeItemExtension.SetAttribute("ExamID", _DicExamID[i]);

                                                    XmlElement SubExam1 = doc1.CreateElement("SubExam");

                                                    XmlElement SubExam2 = doc1.CreateElement("SubExam");

                                                    GradeItemExtension.AppendChild(SubExam1);

                                                    GradeItemExtension.AppendChild(SubExam2);

                                                    SubExam1.SetAttribute("ExtName", "CScore");

                                                    SubExam1.SetAttribute("Permission", "Read");

                                                    SubExam1.SetAttribute("SubName", "讀卡");

                                                    SubExam1.SetAttribute("Type", "Number");

                                                    SubExam2.SetAttribute("ExtName", "PScore");

                                                    SubExam2.SetAttribute("Permission", "Editor");

                                                    SubExam2.SetAttribute("SubName", "試卷");

                                                    SubExam2.SetAttribute("Type", "Number");
                                                }
                                            }

                                            // eleGradeItemExtension 有可能會=null， 代表原本Extensions欄位有其他奇怪的東西(EX: <Abbbbbs><Abbbbb>aaaa</Abbbbb></Abbbbbs>)
                                            else
                                            {
                                                XmlElement extensionsBody = doc1.CreateElement("Extension");

                                                ele.AppendChild(extensionsBody);

                                                extensionsBody.SetAttribute("Name", "GradeItemExtension");

                                                XmlElement GradeItemExtension = doc1.CreateElement("GradeItemExtension");

                                                extensionsBody.AppendChild(GradeItemExtension);

                                                GradeItemExtension.SetAttribute("Calc", "SUM");

                                                GradeItemExtension.SetAttribute("ExamID", _DicExamID[i]);

                                                XmlElement SubExam1 = doc1.CreateElement("SubExam");

                                                XmlElement SubExam2 = doc1.CreateElement("SubExam");

                                                GradeItemExtension.AppendChild(SubExam1);

                                                GradeItemExtension.AppendChild(SubExam2);

                                                SubExam1.SetAttribute("ExtName", "CScore");

                                                SubExam1.SetAttribute("Permission", "Read");

                                                SubExam1.SetAttribute("SubName", "讀卡");

                                                SubExam1.SetAttribute("Type", "Number");

                                                SubExam2.SetAttribute("ExtName", "PScore");

                                                SubExam2.SetAttribute("Permission", "Editor");

                                                SubExam2.SetAttribute("SubName", "試卷");

                                                SubExam2.SetAttribute("Type", "Number");
                                            }
                                        }
                                        // 該子成績沒勾選的話，又如果本來有其資料， 用RemoveChild 整陀Element 移掉
                                        else
                                        {
                                            XmlElement eleGradeItemExtension = ele.SelectSingleNode("Extension[@Name='GradeItemExtension']") as XmlElement;

                                            // 需要加eleGradeItemExtension != null判斷，因為欄位Extension 內可能有其他資料
                                            if (eleGradeItemExtension != null)
                                            {
                                                XmlElement eleGradeItem = eleGradeItemExtension.SelectSingleNode("GradeItemExtension[@ExamID=" + "'" + _DicExamID[i] + "'" + "]") as XmlElement;

                                                if (eleGradeItem != null)
                                                {
                                                    eleGradeItemExtension.RemoveChild(eleGradeItem);
                                                }
                                            }
                                        }
                                }

                                // 因為ele 本來有東西，所以更新後一樣傳本體   記得要下SQL 內的String 要用 ' ' 包起來ㄚㄚ~
                                body = "'" + ele.OuterXml + "'";
                            }

                            // 組合最終要上傳的SQL 指令
                            String sql = @"UPDATE course SET extensions =" + body + " WHERE id=" + SPCR.ID;

                            // 使用 UpdateHelper 執行SQL
                            uh.Execute(sql);
                        }
                    }
                }
            }

            // 更新完畢後關閉
            this.Close();
        }



        // 取消
        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
