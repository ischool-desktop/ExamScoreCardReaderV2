using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using System.IO;
using K12.Data;
using ExamScoreCardReaderV2.Model;
using ExamScoreCardReaderV2.Validation;
using ExamScoreCardReaderV2.Validation.RecordValidators;
using System.Xml;

namespace ExamScoreCardReaderV2
{
    public partial class ImportStartupForm : BaseForm
    {

        private BackgroundWorker _worker;
        private BackgroundWorker _upload;
        private BackgroundWorker _warn;

        private int SchoolYear { get; set; }
        private int Semester { get; set; }

        private int StudentNumberMax { get; set; }

        private DataValidator<RawData> _rawDataValidator;
        private DataValidator<DataRecord> _dataRecordValidator;

        private List<FileInfo> _files;
        private List<SCETakeRecord> _addScoreList;
        private List<SCETakeRecord> _existedScoreList;
        private Dictionary<string, SCETakeRecord> AddScoreDic;
        private Dictionary<string, string> examDict;

        private Dictionary<string, SCETakeRecord> dicAddScore_Raw;
        private Dictionary<string, SCETakeRecord> dicUpdateScore;
        private List<SCETakeRecord> _addScoreList_Upload;

        // 2016/8/4 穎驊新增，紀錄該課程ID是否有子成績
        public Dictionary<String, bool> CoursesWithSubScore = new Dictionary<string, bool>();


        StringBuilder sbLog { get; set; }
        /// <summary>
        /// 儲存畫面上學號長度
        /// </summary>
        K12.Data.Configuration.ConfigData cd;

        private string _ExamScoreReaderConfig = "高中讀卡系統設定檔";

        private string _StudentDotIsClear = "是否移除小數點後內容";
        private string _StudentNumberLenghtName = "國中匯入讀卡學號長度";

        //高中系統努力程度已無使用
        //private EffortMapper _effortMapper;

        int counter = 0; //上傳成績時，算筆數用的。

        /// <summary>
        /// 載入學號長度值
        /// </summary>
        private void LoadConfigData()
        {
            cd = School.Configuration[_ExamScoreReaderConfig];

            int val1 = 0;


            if (int.TryParse(cd[_StudentNumberLenghtName], out val1))
            {
                Global.StudentNumberLenght = val1;
            }

            if (Global.StudentNumberLenght < 5)
            {
                Global.StudentNumberLenght = 5;
            }
            else
            {

            }
            // 學號顯示預設為 上一次的選得碼數，由於我們的第一筆學號為從5++ 到10
            //以5碼為例，他是List 的第5 - 5 =第0項，也就是第一筆資料~
            cboStudentNumberMax.SelectedIndex = Global.StudentNumberLenght - 5;


            bool val2 = false;
            Global.StudentDocRemove = val2; //預設是不移除
            if (bool.TryParse(cd[_StudentDotIsClear], out val2))
                checkBoxX1.Checked = val2;

        }


        /// <summary>
        /// 儲存學號長度值
        /// </summary>
        private void SaveConfigData()
        {
            int ii = Global.StudentNumberLenght;

            Global.StudentNumberLenght = StudentNumberMax;
            cd[_StudentNumberLenghtName] = StudentNumberMax.ToString();
            cd.Save();
        }

        public ImportStartupForm()
        {
            InitializeComponent();
            InitializeSemesters();

            //_effortMapper = new EffortMapper();

            // 載入預設儲存值
            LoadConfigData();

            _worker = new BackgroundWorker();
            _worker.WorkerReportsProgress = true;
            _worker.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            {
                lblMessage.Text = "" + e.UserState;
            };
            _worker.DoWork += delegate(object sender, DoWorkEventArgs e)
            {
                #region Worker DoWork
                _worker.ReportProgress(0, "訊息：檢查讀卡文字格式…");

                #region 檢查文字檔
                ValidateTextFiles vtf = new ValidateTextFiles(StudentNumberMax);
                ValidateTextResult vtResult = vtf.CheckFormat(_files);
                if (vtResult.Error)
                {
                    e.Result = vtResult;
                    return;
                }
                #endregion

                //文字檔轉 RawData
                RawDataCollection rdCollection = new RawDataCollection();
                rdCollection.ConvertFromFiles(_files);

                //RawData 轉 DataRecord
                DataRecordCollection drCollection = new DataRecordCollection();
                drCollection.ConvertFromRawData(rdCollection);

                _rawDataValidator = new DataValidator<RawData>();
                _dataRecordValidator = new DataValidator<DataRecord>();

                #region 取得驗證需要的資料
                Course.RemoveAll();
                _worker.ReportProgress(5, "訊息：取得學生資料…");

                List<StudentObj> studentList = GetInSchoolStudents();

                List<string> s_ids = new List<string>();
                Dictionary<string, List<string>> studentNumberToStudentIDs = new Dictionary<string, List<string>>();
                foreach (StudentObj student in studentList)
                {
                    string sn = student.StudentNumber;// SCValidatorCreator.GetStudentNumberFormat(student.StudentNumber);
                    if (!studentNumberToStudentIDs.ContainsKey(sn))
                        studentNumberToStudentIDs.Add(sn, new List<string>());
                    studentNumberToStudentIDs[sn].Add(student.StudentID);
                }

                foreach (string each in studentNumberToStudentIDs.Keys)
                {
                    if (studentNumberToStudentIDs[each].Count > 1)
                    {
                        //學號重覆
                    }
                }

                foreach (var dr in drCollection)
                {
                    if (studentNumberToStudentIDs.ContainsKey(dr.StudentNumber))
                        s_ids.AddRange(studentNumberToStudentIDs[dr.StudentNumber]);
                }

                studentList.Clear();

                _worker.ReportProgress(10, "訊息：取得課程資料…");
                List<CourseRecord> courseList = Course.SelectBySchoolYearAndSemester(SchoolYear, Semester);
                List<AEIncludeRecord> aeList = AEInclude.SelectAll();

                //List<SCAttendRecord> scaList = SCAttend.SelectAll();
                var c_ids = from course in courseList select course.ID;
                _worker.ReportProgress(15, "訊息：取得修課資料…");
                //List<SCAttendRecord> scaList2 = SCAttend.SelectByStudentIDAndCourseID(s_ids, c_ids.ToList<string>());
                List<SCAttendRecord> scaList = new List<SCAttendRecord>();
                FunctionSpliter<string, SCAttendRecord> spliter = new FunctionSpliter<string, SCAttendRecord>(300, 3);
                spliter.Function = delegate(List<string> part)
                {
                    return SCAttend.Select(part, c_ids.ToList<string>(), null, SchoolYear.ToString(), Semester.ToString());
                };
                scaList = spliter.Execute(s_ids);

                _worker.ReportProgress(20, "訊息：取得試別資料…");
                List<ExamRecord> examList = Exam.SelectAll();
                #endregion

                #region 註冊驗證
                _worker.ReportProgress(30, "訊息：載入驗證規則…");
                _rawDataValidator.Register(new SubjectCodeValidator());
                _rawDataValidator.Register(new ClassCodeValidator());
                _rawDataValidator.Register(new ExamCodeValidator());

                SCValidatorCreator scCreator = new SCValidatorCreator(Student.SelectByIDs(s_ids), courseList, scaList);
                _dataRecordValidator.Register(scCreator.CreateStudentValidator());
                _dataRecordValidator.Register(new ExamValidator(examList));
                _dataRecordValidator.Register(scCreator.CreateSCAttendValidator());
                _dataRecordValidator.Register(new CourseExamValidator(scCreator.StudentCourseInfo, aeList, examList));
                #endregion

                #region 進行驗證
                _worker.ReportProgress(45, "訊息：進行驗證中…");
                List<string> msgList = new List<string>();

                foreach (RawData rawData in rdCollection)
                {
                    List<string> msgs = _rawDataValidator.Validate(rawData);
                    msgList.AddRange(msgs);
                }
                if (msgList.Count > 0)
                {
                    e.Result = msgList;
                    return;
                }

                foreach (DataRecord dataRecord in drCollection)
                {
                    List<string> msgs = _dataRecordValidator.Validate(dataRecord);
                    msgList.AddRange(msgs);
                }
                if (msgList.Count > 0)
                {
                    e.Result = msgList;
                    return;
                }
                #endregion

                #region 取得學生的評量成績

                _worker.ReportProgress(65, "訊息：取得學生評量成績…");

                _existedScoreList.Clear();
                _addScoreList.Clear();
                AddScoreDic.Clear();

                //var student_ids = from student in scCreator.StudentNumberDictionary.Values select student.ID;
                //List<string> course_ids = scCreator.AttendCourseIDs;

                var scaIDs = from sca in scaList select sca.ID;

                Dictionary<string, SCETakeRecord> sceList = new Dictionary<string, SCETakeRecord>();
                FunctionSpliter<string, SCETakeRecord> spliterSCE = new FunctionSpliter<string, SCETakeRecord>(300, 3);
                spliterSCE.Function = delegate(List<string> part)
                {
                    return SCETake.Select(null, null, null, null, part);
                };
                foreach (SCETakeRecord sce in spliterSCE.Execute(scaIDs.ToList()))
                {
                    string key = GetCombineKey(sce.RefStudentID, sce.RefCourseID, sce.RefExamID);
                    if (!sceList.ContainsKey(key))
                        sceList.Add(key, sce);
                }

                Dictionary<string, ExamRecord> examTable = new Dictionary<string, ExamRecord>();
                Dictionary<string, SCAttendRecord> scaTable = new Dictionary<string, SCAttendRecord>();

                foreach (ExamRecord exam in examList)
                    if (!examTable.ContainsKey(exam.Name))
                        examTable.Add(exam.Name, exam);

                foreach (SCAttendRecord sca in scaList)
                {
                    string key = GetCombineKey(sca.RefStudentID, sca.RefCourseID);
                    if (!scaTable.ContainsKey(key))
                        scaTable.Add(key, sca);
                }

                _worker.ReportProgress(80, "訊息：成績資料建立…");

                foreach (DataRecord dr in drCollection)
                {
                    StudentRecord student = student = scCreator.StudentNumberDictionary[dr.StudentNumber];
                    ExamRecord exam = examTable[dr.Exam];
                    List<CourseRecord> courses = new List<CourseRecord>();
                    foreach (CourseRecord course in scCreator.StudentCourseInfo.GetCourses(dr.StudentNumber))
                    {
                        if (dr.Subjects.Contains(course.Subject))
                            courses.Add(course);
                    }

                    foreach (CourseRecord course in courses)
                    {
                        string key = GetCombineKey(student.ID, course.ID, exam.ID);

                        if (sceList.ContainsKey(key))
                            _existedScoreList.Add(sceList[key]);

                        SCETakeRecord sh = new SCETakeRecord();
                        sh.RefCourseID = course.ID;
                        sh.RefExamID = exam.ID;
                        sh.RefSCAttendID = scaTable[GetCombineKey(student.ID, course.ID)].ID;
                        sh.RefStudentID = student.ID;

                        String CoursesWithSubScore_Key = sh.RefCourseID + "_" + sh.RefExamID;

                        if (!CoursesWithSubScore.ContainsKey(CoursesWithSubScore_Key))
                        {
                            string strSQL_course = "SELECT * FROM course WHERE id =" + sh.RefCourseID;

                            FISCA.Data.QueryHelper qh = new FISCA.Data.QueryHelper();

                            System.Data.DataTable dt_course = qh.Select(strSQL_course);


                            XmlDocument doc1 = new XmlDocument();

                            string strXml = "" + dt_course.Rows[0]["Extensions"];

                            if (strXml != "")
                            {
                                doc1.LoadXml(strXml);
                            }

                            //指定 ele 為 Xmldocument doc1 的Element
                            XmlElement ele = doc1.DocumentElement;

                            if (ele != null)
                            {
                                XmlElement eleGradeItemExtension = ele.SelectSingleNode("Extension[@Name='GradeItemExtension']") as XmlElement;

                                // eleGradeItemExtension 有可能會=null， 代表原本Extensions欄位有其他奇怪的東西(EX: <Abbbbbs><Abbbbb>aaaa</Abbbbb></Abbbbbs>) 而通過第一關ele != null 的檢驗
                                if (eleGradeItemExtension != null)
                                {
                                    XmlElement eleGradeItem = eleGradeItemExtension.SelectSingleNode("GradeItemExtension[@ExamID=" + "'" + sh.RefExamID + "'" + "]") as XmlElement;

                                    // 
                                    if (eleGradeItem != null)
                                    {
                                        CoursesWithSubScore.Add(CoursesWithSubScore_Key, true);
                                    }
                                    //eleGradeItem  = null 代表你Extensions 裡面本科本次考試沒有子成績的設定，判定該科目本次考試沒有子成績
                                    else
                                    {
                                        CoursesWithSubScore.Add(CoursesWithSubScore_Key, false);
                                    }
                                }
                                //eleGradeItemExtension  = null 代表你Extensions 裡面沒有任何一科子成績的設定，判定該科目本次考試沒有子成績
                                else
                                {
                                    CoursesWithSubScore.Add(CoursesWithSubScore_Key, false);
                                }
                            }
                            //ele  = null 代表你Extensions 裡面甚麼都沒有，判定該科目本次考試沒有子成績
                            else
                            {
                                CoursesWithSubScore.Add(CoursesWithSubScore_Key, false);
                            }
                        }


                        sh.SetScore(dr.Score, Global.StudentDocRemove, CoursesWithSubScore);


                        ////轉型Double再轉回decimal,可去掉小數點後的0
                        //double reScore = (double)dr.Score;
                        //decimal Score = decimal.Parse(reScore.ToString());

                        //if (Global.StudentDocRemove)
                        //{
                        //    string qq = Score.ToString();
                        //    if (qq.Contains("."))
                        //    {
                        //        string[] kk = qq.Split('.');
                        //        sh.Score = decimal.Parse(kk[0]);
                        //    }
                        //    else
                        //    {
                        //        sh.Score = decimal.Parse(Score.ToString());
                        //    }
                        //}
                        //else
                        //{
                        //    sh.Score = decimal.Parse(Score.ToString());
                        //}

                        //sceNew.Effort = _effortMapper.GetCodeByScore(dr.Score);

                        //是否有重覆的學生,課程,評量
                        if (!AddScoreDic.ContainsKey(sh.RefStudentID + "_" + course.ID + "_" + exam.ID))
                        {
                            _addScoreList.Add(sh);
                            AddScoreDic.Add(sh.RefStudentID + "_" + course.ID + "_" + exam.ID, sh);
                        }
                    }
                }
                #endregion
                _worker.ReportProgress(100, "訊息：背景作業完成…");
                e.Result = null;
                #endregion
            };
            _worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                #region Worker Completed
                if (e.Error == null && e.Result == null)
                {
                    if (!_upload.IsBusy)
                    {
                        //如果學生身上已有成績，則提醒使用者
                        if (_existedScoreList.Count > 0)
                        {
                            _warn.RunWorkerAsync();
                        }
                        else
                        {
                            lblMessage.Text = "訊息：成績上傳中…";
                            FISCA.Presentation.MotherForm.SetStatusBarMessage("成績上傳中…", 0);
                            counter = 0;
                            _upload.RunWorkerAsync();
                        }
                    }
                }
                else
                {
                    ControlEnable = true;

                    if (e.Error != null)
                    {
                        MsgBox.Show("匯入失敗。" + e.Error.Message);
                        SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);

                    }
                    else if (e.Result != null && e.Result is ValidateTextResult)
                    {
                        ValidateTextResult result = e.Result as ValidateTextResult;
                        ValidationErrorViewer viewer = new ValidationErrorViewer();
                        viewer.SetTextFileError(result.LineIndexes, result.ErrorFormatLineIndexes, result.DuplicateLineIndexes);
                        viewer.ShowDialog();
                    }
                    else if (e.Result != null && e.Result is List<string>)
                    {
                        ValidationErrorViewer viewer = new ValidationErrorViewer();
                        viewer.SetErrorLines(e.Result as List<string>);
                        viewer.ShowDialog();
                    }
                }
                #endregion
            };

            _upload = new BackgroundWorker();
            _upload.WorkerReportsProgress = true;
            _upload.ProgressChanged += new ProgressChangedEventHandler(_upload_ProgressChanged);
            _upload.DoWork += new DoWorkEventHandler(_upload_DoWork);


            _upload.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_upload_RunWorkerCompleted);

            _warn = new BackgroundWorker();
            _warn.WorkerReportsProgress = true;
            _warn.DoWork += delegate(object sender, DoWorkEventArgs e)
            {
                _warn.ReportProgress(0, "產生警告訊息...");

                examDict = new Dictionary<string, string>();
                foreach (ExamRecord exam in Exam.SelectAll())
                {
                    if (!examDict.ContainsKey(exam.ID))
                        examDict.Add(exam.ID, exam.Name);
                }

                WarningForm form = new WarningForm();
                int count = 0;
                foreach (SCETakeRecord sce in _existedScoreList)
                {
                    // 當成績資料是空值跳過
                    //if (sce.GetScore().HasValue == false && sce.Effort.HasValue == false && string.IsNullOrEmpty(sce.Text))
                    //if (sce.GetScore() == null && string.IsNullOrEmpty(sce.Text))
                    //   continue;

                    count++;


                    StudentRecord student = Student.SelectByID(sce.RefStudentID);
                    CourseRecord course = Course.SelectByID(sce.RefCourseID);
                    string exam = (examDict.ContainsKey(sce.RefExamID) ? examDict[sce.RefExamID] : "<未知的試別>");

                    string s = "";
                    if (student.Class != null) s += student.Class.Name;
                    if (!string.IsNullOrEmpty("" + student.SeatNo)) s += " " + student.SeatNo + "號";
                    if (!string.IsNullOrEmpty(student.StudentNumber)) s += " (" + student.StudentNumber + ")";
                    s += " " + student.Name;

                    string scoreName = sce.RefStudentID + "_" + sce.RefCourseID + "_" + sce.RefExamID;

                    if (AddScoreDic.ContainsKey(scoreName))
                    {
                        form.AddMessage(student.ID, s, string.Format("學生在「{0}」課程「{1}」中已有讀卡成績「{2}」將修改為「{3}」。", course.Name, exam, sce.GetScore(), AddScoreDic[scoreName].GetScore()));
                    }
                    else
                    {
                        form.AddMessage(student.ID, s, string.Format("學生在「{0}」課程「{1}」中已有讀卡成績「{2}」。", course.Name, exam, sce.GetScore()));
                    }

                    _warn.ReportProgress((int)(count * 100 / _existedScoreList.Count), "產生警告訊息...");
                }

                e.Result = form;
            };
            _warn.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                WarningForm form = e.Result as WarningForm;

                if (form.ShowDialog() == DialogResult.OK)
                {
                    lblMessage.Text = "訊息：成績上傳中…";
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("成績上傳中…", 0);
                    counter = 0;
                    _upload.RunWorkerAsync();
                }
                else
                {
                    this.DialogResult = DialogResult.Cancel;
                }
            };
            _warn.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("" + e.UserState, e.ProgressPercentage);
            };

            _files = new List<FileInfo>();
            _addScoreList = new List<SCETakeRecord>();
            _existedScoreList = new List<SCETakeRecord>();
            AddScoreDic = new Dictionary<string, SCETakeRecord>();
            examDict = new Dictionary<string, string>();
        }

        void _upload_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("成績上傳中…", (int)(100f * counter / (double)_addScoreList.Count));
        }

        // 上傳成績完成
        void _upload_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ControlEnable = true;

            if (!e.Cancelled)
            {
                if (e.Error == null)
                {
                    string msg = "";

                    if (e.Result != null)
                        msg = e.Result.ToString();

                    FISCA.Presentation.MotherForm.SetStatusBarMessage("匯入評量讀卡成績已完成!!共" + msg + "筆");

                    LogViewFrom lv = new LogViewFrom(sbLog.ToString());
                    lv.ShowDialog();
                }
            }
            else
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("匯入作業已中止!!");
                MsgBox.Show("匯入作業已中止!!");
            }
        }


        // 上傳
        void _upload_DoWork(object sender, DoWorkEventArgs e)
        {

            Dictionary<string, SCETakeRecord> dicExistedOri = new Dictionary<string, SCETakeRecord>();
            foreach (SCETakeRecord sce in _existedScoreList)
            {
                if (!dicExistedOri.ContainsKey(sce.RefStudentID + "_" + sce.RefCourseID))
                {
                    dicExistedOri.Add(sce.RefStudentID + "_" + sce.RefCourseID + "_" + sce.RefExamID, new SCETakeRecord()
                    {
                        RefSCAttendID = sce.RefSCAttendID,
                        RefStudentID = sce.RefStudentID,
                        RefCourseID = sce.RefCourseID,
                        RefExamID = sce.RefExamID
                    });
                }
            }

            dicAddScore_Raw = new Dictionary<string, SCETakeRecord>();

            dicUpdateScore = new Dictionary<string, SCETakeRecord>();

            _addScoreList_Upload = new List<SCETakeRecord>();

            // 傳送與回傳筆數
            int SendCount = 0;
            // 刪除舊資料
            SendCount = _existedScoreList.Count;

            // 取得 del id
            List<string> delIDList = _existedScoreList.Select(x => x.ID).ToList();

            //2016/5/6 穎驊改寫，因應新竹國中有一期定期成績綁平時分數的現象，故不可以每次用原作法都把SCETRecord檔案刪掉新增
            // 所以需要動用到Dictionary 比對key、value，使用Update更新資料，目前接手是直接使用既有的_deleteScoreList的項目
            // _deleteScoreList的每一項目 都是資料有重覆的，要進行Update，故把他們加到UpdateScoreDic，比對後上傳更新Update，其餘
            //的項目則使用Insert 來新增

            foreach (var y in _addScoreList)
            {
                var key = y.RefStudentID + "_" + y.RefCourseID + "_" + y.RefExamID;
                dicAddScore_Raw.Add(key, y);
                #region 檢查有沒有子成績
                String CoursesWithSubScore_Key = dicAddScore_Raw[key].RefCourseID + "_" + dicAddScore_Raw[key].RefExamID;

                if (!CoursesWithSubScore.ContainsKey(CoursesWithSubScore_Key))
                {
                    string strSQL_course = "SELECT * FROM course WHERE id =" + dicAddScore_Raw[key].RefCourseID;

                    FISCA.Data.QueryHelper qh = new FISCA.Data.QueryHelper();

                    System.Data.DataTable dt_course = qh.Select(strSQL_course);


                    XmlDocument doc1 = new XmlDocument();

                    string strXml = "" + dt_course.Rows[0]["Extensions"];

                    if (strXml != "")
                    {
                        doc1.LoadXml(strXml);
                    }

                    //指定 ele 為 Xmldocument doc1 的Element
                    XmlElement ele = doc1.DocumentElement;

                    if (ele != null)
                    {
                        XmlElement eleGradeItemExtension = ele.SelectSingleNode("Extension[@Name='GradeItemExtension']") as XmlElement;

                        // eleGradeItemExtension 有可能會=null， 代表原本Extensions欄位有其他奇怪的東西(EX: <Abbbbbs><Abbbbb>aaaa</Abbbbb></Abbbbbs>) 而通過第一關ele != null 的檢驗
                        if (eleGradeItemExtension != null)
                        {
                            XmlElement eleGradeItem = eleGradeItemExtension.SelectSingleNode("GradeItemExtension[@ExamID=" + "'" + dicAddScore_Raw[key].RefExamID + "'" + "]") as XmlElement;

                            // 
                            if (eleGradeItem != null)
                            {
                                CoursesWithSubScore.Add(CoursesWithSubScore_Key, true);
                            }
                            //eleGradeItem  = null 代表你Extensions 裡面本科本次考試沒有子成績的設定，判定該科目本次考試沒有子成績
                            else
                            {
                                CoursesWithSubScore.Add(CoursesWithSubScore_Key, false);
                            }
                        }
                        //eleGradeItemExtension  = null 代表你Extensions 裡面沒有任何一科子成績的設定，判定該科目本次考試沒有子成績
                        else
                        {
                            CoursesWithSubScore.Add(CoursesWithSubScore_Key, false);
                        }
                    }
                    //ele  = null 代表你Extensions 裡面甚麼都沒有，判定該科目本次考試沒有子成績
                    else
                    {
                        CoursesWithSubScore.Add(CoursesWithSubScore_Key, false);
                    }
                }
                #endregion
            }

            foreach (var x in _existedScoreList)
            {
                dicUpdateScore.Add(x.RefStudentID + "_" + x.RefCourseID + "_" + x.RefExamID, x);
            }

            foreach (var r in dicAddScore_Raw.Keys)
            {

                if (dicUpdateScore.ContainsKey(r))
                {
                    dicExistedOri[r].SetScore(dicUpdateScore[r].GetScore(), true, CoursesWithSubScore);
                    dicUpdateScore[r].SetScore(dicAddScore_Raw[r].GetScore(), true, CoursesWithSubScore);
                }
                else
                {
                    _addScoreList_Upload.Add(dicAddScore_Raw[r]);
                }
            }

            // 執行          
            //try
            //{
            //    SCETake.Update(_deleteScoreList);
            //}
            //catch (Exception ex)
            //{
            //    e.Result = ex.Message;
            //    e.Cancel = true;
            //}
            //    RspCount = SCETake.SelectByIDs(delIDList).Count;
            //// 刪除未完成
            //    if (RspCount > 0)
            //        e.Cancel = true;

            try
            {
                #region 分筆更新
                {
                    Dictionary<int, List<SCETakeRecord>> batchDict = new Dictionary<int, List<SCETakeRecord>>();
                    int bn = 150;
                    int n1 = (int)(dicUpdateScore.Count / bn);

                    if ((dicUpdateScore.Count % bn) != 0)
                        n1++;

                    for (int i = 0; i <= n1; i++)
                        batchDict.Add(i, new List<SCETakeRecord>());


                    if (dicUpdateScore.Count > 0)
                    {
                        int idx = 0, count = 1;
                        // 分批
                        foreach (SCETakeRecord rec in dicUpdateScore.Values)
                        {
                            // 100 分一批
                            if ((count % bn) == 0)
                                idx++;

                            batchDict[idx].Add(rec);
                            count++;
                        }
                    }


                    //上傳資料
                    foreach (KeyValuePair<int, List<SCETakeRecord>> data in batchDict)
                    {
                        SendCount = 0;
                        if (data.Value.Count > 0)
                        {
                            SendCount = data.Value.Count;
                            try
                            {
                                SCETake.Update(data.Value);

                            }
                            catch (Exception ex)
                            {
                                e.Cancel = true;
                                e.Result = ex.Message;
                            }

                            counter += SendCount;

                        }
                    }
                }
                #endregion

                #region 分筆新增
                {
                    Dictionary<int, List<SCETakeRecord>> batchDict = new Dictionary<int, List<SCETakeRecord>>();
                    int bn = 150;
                    int n1 = (int)(_addScoreList_Upload.Count / bn);

                    if ((_addScoreList_Upload.Count % bn) != 0)
                        n1++;

                    for (int i = 0; i <= n1; i++)
                        batchDict.Add(i, new List<SCETakeRecord>());


                    if (_addScoreList_Upload.Count > 0)
                    {
                        int idx = 0, count = 1;
                        // 分批
                        foreach (SCETakeRecord rec in _addScoreList_Upload)
                        {
                            // 100 分一批
                            if ((count % bn) == 0)
                                idx++;

                            batchDict[idx].Add(rec);
                            count++;
                        }
                    }


                    //上傳資料
                    foreach (KeyValuePair<int, List<SCETakeRecord>> data in batchDict)
                    {
                        SendCount = 0;
                        if (data.Value.Count > 0)
                        {
                            SendCount = data.Value.Count;
                            try
                            {
                                SCETake.Insert(data.Value);

                            }
                            catch (Exception ex)
                            {
                                e.Cancel = true;
                                e.Result = ex.Message;
                            }

                            counter += SendCount;

                        }
                    }
                }
                #endregion

                SetLog(dicExistedOri);
                e.Result = counter;
            }
            catch (Exception ex)
            {
                e.Result = ex.Message;
                e.Cancel = true;

            }
        }

        private void SetLog(Dictionary<string, SCETakeRecord> delExistedOri)
        {
            sbLog = new StringBuilder();

            foreach (SCETakeRecord sce in _addScoreList)
            {
                StudentRecord student = Student.SelectByID(sce.RefStudentID);
                CourseRecord course = Course.SelectByID(sce.RefCourseID);
                string exam = (examDict.ContainsKey(sce.RefExamID) ? examDict[sce.RefExamID] : "<未知的試別>");

                if (delExistedOri.ContainsKey(sce.RefStudentID + "_" + sce.RefCourseID + "_" + sce.RefExamID)
                    && delExistedOri[sce.RefStudentID + "_" + sce.RefCourseID + "_" + sce.RefExamID].GetScore() != null)
                {
                    string classname = student.Class != null ? student.Class.Name : "";
                    string seatno = student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "";
                    sbLog.AppendLine(string.Format("班級「{0}」座號「{1}」姓名「{2}」在試別「{3}」課程「{4}」將讀卡成績「{5}」修改為「{6}」。", classname, seatno, student.Name, course.Name, exam, delExistedOri[sce.RefStudentID + "_" + sce.RefCourseID + "_" + sce.RefExamID].GetScore(), sce.GetScore()));

                }
                else
                {
                    string classname = student.Class != null ? student.Class.Name : "";
                    string seatno = student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "";
                    sbLog.AppendLine(string.Format("班級「{0}」座號「{1}」姓名「{2}」在「{3}」課程「{4}」新增讀卡成績「{5}」。", classname, seatno, student.Name, course.Name, exam, sce.GetScore()));
                }
            }

            FISCA.LogAgent.ApplicationLog.Log("讀卡系統", "評量成績", sbLog.ToString());
        }

        private List<StudentObj> GetInSchoolStudents()
        {
            List<StudentObj> list = new List<StudentObj>();
            StringBuilder sb = new StringBuilder();
            sb.Append("select id,student_number from student where status in (1) and student_number is not null");
            DataTable dt = tool._Q.Select(sb.ToString());

            foreach (DataRow row in dt.Rows)
            {
                StudentObj obj = new StudentObj(row);

                list.Add(obj);
            }
            return list;
        }

        private string GetCombineKey(string s1, string s2, string s3)
        {
            return s1 + "_" + s2 + "_" + s3;
        }

        private string GetCombineKey(string s1, string s2)
        {
            return s1 + "_" + s2;
        }

        private void InitializeSemesters()
        {
            try
            {
                for (int i = -2; i <= 2; i++)
                {
                    cboSchoolYear.Items.Add(int.Parse(School.DefaultSchoolYear) + i);
                }
                cboSemester.Items.Add(1);
                cboSemester.Items.Add(2);

                cboSchoolYear.SelectedIndex = 2;
                cboSemester.SelectedIndex = int.Parse(School.DefaultSemester) - 1;



                // 加入學號選擇碼數，目前為5~10碼
                cboStudentNumberMax.Items.Add(5);
                cboStudentNumberMax.Items.Add(6);
                cboStudentNumberMax.Items.Add(7);
                cboStudentNumberMax.Items.Add(8);
                cboStudentNumberMax.Items.Add(9);
                cboStudentNumberMax.Items.Add(10);




            }
            catch { }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;
            ofd.Filter = "純文字文件(*.txt)|*.txt";
            ofd.Multiselect = true;
            ofd.Title = "開啟檔案";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                _files.Clear();
                StringBuilder builder = new StringBuilder("");
                foreach (var file in ofd.FileNames)
                {
                    FileInfo fileInfo = new FileInfo(file);
                    _files.Add(fileInfo);
                    builder.Append(fileInfo.Name + ", ");
                }
                string fileString = builder.ToString();
                if (fileString.EndsWith(", ")) fileString = fileString.Substring(0, fileString.Length - 2);
                txtFiles.Text = fileString;
            }
        }

        private bool ControlEnable
        {
            set
            {
                foreach (Control ctrl in this.Controls)
                    ctrl.Enabled = value;

                pic.Enabled = lblMessage.Enabled = !value;
                pic.Visible = lblMessage.Visible = !value;
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (cboSchoolYear.SelectedItem == null) return;
            if (cboSemester.SelectedItem == null) return;
            if (cboStudentNumberMax.SelectedItem == null) return;
            if (_files.Count <= 0) return;

            ControlEnable = false;

            // 儲存設定值
            SaveConfigData();

            if (!_worker.IsBusy)
                _worker.RunWorkerAsync();
        }

        private void cboSchoolYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboSchoolYear.SelectedItem != null)
                SchoolYear = (int)cboSchoolYear.SelectedItem;
        }

        private void cboSemester_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboSemester.SelectedItem != null)
                Semester = (int)cboSemester.SelectedItem;
        }

        private void ImportStartupForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ImportStartupForm_Load(object sender, EventArgs e)
        {

        }


        // 2016/5/9 穎驊 新增，使得每間學校可以依照學號位數的不同在表上動態調整5~10碼

        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (cboStudentNumberMax.SelectedItem != null)
                StudentNumberMax = (int)cboStudentNumberMax.SelectedItem;
        }
    }

    internal static class Ext
    {

        public static void SetScore(this SCETakeRecord target, decimal? score, bool trimEnd, Dictionary<String, bool> CoursesWithSubScore)
        {
            XmlElement element = target.ToXML();

            //  取得當前XmlElement 的Document，超重要，如果另外New一個Document ，兩個不同的Document 產生的Element 不能互相穿插
            XmlDocument doc = element.OwnerDocument;

            String CoursesWithSubScore_Key = target.RefCourseID + "_" + target.RefExamID;

            //沒有子成績
            if (!CoursesWithSubScore[CoursesWithSubScore_Key])
            {
                element.SelectSingleNode("Score").InnerText = K12.Data.Decimal.GetString(score);

                XmlElement Extension = element.SelectSingleNode("Extension") as XmlElement;


                XmlElement Extension_inner = Extension.SelectSingleNode("Extension") as XmlElement;

                if (Extension_inner != null)
                {

                    XmlElement eleCScore = Extension_inner.SelectSingleNode("CScore") as XmlElement;

                    if (eleCScore == null)
                    {
                        XmlElement CScore = doc.CreateElement("CScore");
                        Extension_inner.AppendChild(CScore);
                        CScore.InnerText = K12.Data.Decimal.GetString(score);
                    }
                    else
                    {
                        eleCScore.InnerText = K12.Data.Decimal.GetString(score);
                    }

                    XmlElement elePScore = Extension_inner.SelectSingleNode("PScore") as XmlElement;

                    if (elePScore == null)
                    {
                        XmlElement PScore = doc.CreateElement("PScore");
                        Extension_inner.AppendChild(PScore);
                        PScore.InnerText = K12.Data.Decimal.GetString(0);
                    }

                    XmlElement eleScore = Extension_inner.SelectSingleNode("Score") as XmlElement;
                    if (eleScore == null)
                    {
                        XmlElement Score = doc.CreateElement("Score");
                        Extension_inner.AppendChild(Score);
                        Score.InnerText = K12.Data.Decimal.GetString(score);
                    }
                    else
                    {
                        eleScore.InnerText = K12.Data.Decimal.GetString(score);
                    }


                    element.SelectSingleNode("Score").InnerText = Extension_inner.SelectSingleNode("Score").InnerText;

                }
                else
                {

                    XmlElement Extension_innerCreate = doc.CreateElement("Extension");

                    Extension.AppendChild(Extension_innerCreate);


                    XmlElement CScore = doc.CreateElement("CScore");
                    Extension_innerCreate.AppendChild(CScore);
                    CScore.InnerText = K12.Data.Decimal.GetString(score);


                    XmlElement PScore = doc.CreateElement("PScore");
                    Extension_innerCreate.AppendChild(PScore);
                    PScore.InnerText = K12.Data.Decimal.GetString(0);

                    XmlElement Score = doc.CreateElement("Score");
                    Extension_innerCreate.AppendChild(Score);
                    Score.InnerText = K12.Data.Decimal.GetString(score);

                    element.SelectSingleNode("Score").InnerText = Extension_inner.SelectSingleNode("Score").InnerText;


                }

                //Extension.SelectSingleNode("Score").InnerText = K12.Data.Decimal.GetString(score);

                //Extension.SelectSingleNode("CScore").InnerText = K12.Data.Decimal.GetString(score);

                //Extension.SelectSingleNode("PScore").InnerText = "0";

            }
            // 有子成績
            else
            {

                XmlElement Extension = element.SelectSingleNode("Extension") as XmlElement;

                XmlElement Extension_inner = Extension.SelectSingleNode("Extension") as XmlElement;

                if (Extension_inner != null)
                {

                    XmlElement eleCScore = Extension_inner.SelectSingleNode("CScore") as XmlElement;

                    if (eleCScore == null)
                    {
                        XmlElement CScore = doc.CreateElement("CScore");
                        Extension_inner.AppendChild(CScore);
                        CScore.InnerText = K12.Data.Decimal.GetString(score);
                    }
                    else
                    {
                        eleCScore.InnerText = K12.Data.Decimal.GetString(score);
                    }

                    XmlElement elePScore = Extension_inner.SelectSingleNode("PScore") as XmlElement;

                    if (elePScore == null)
                    {
                        XmlElement PScore = doc.CreateElement("PScore");
                        Extension_inner.AppendChild(PScore);
                        //PScore.InnerText = K12.Data.Decimal.GetString(0);
                    }

                    //擴充欄位值
                    XmlElement eleScore = Extension_inner.SelectSingleNode("Score") as XmlElement;
                    if (eleScore == null)
                    {
                        eleScore = doc.CreateElement("Score");
                        Extension_inner.AppendChild(eleScore);
                    }
                    decimal pScore = 0;
                    if (!decimal.TryParse(Extension_inner.SelectSingleNode("PScore").InnerText, out pScore) && Extension_inner.SelectSingleNode("PScore").InnerText == "缺")
                        eleScore.InnerText = "缺";
                    else
                        eleScore.InnerText = K12.Data.Decimal.GetString(score + pScore);

                    //實體欄位值
                    if (eleScore.InnerText == "缺")
                        element.SelectSingleNode("Score").InnerText = "-1";
                    else
                        element.SelectSingleNode("Score").InnerText = Extension_inner.SelectSingleNode("Score").InnerText;

                }
                else
                {
                    XmlElement Extension_innerCreate = doc.CreateElement("Extension");

                    Extension.AppendChild(Extension_innerCreate);


                    XmlElement CScore = doc.CreateElement("CScore");
                    Extension_innerCreate.AppendChild(CScore);
                    CScore.InnerText = K12.Data.Decimal.GetString(score);


                    XmlElement PScore = doc.CreateElement("PScore");
                    Extension_innerCreate.AppendChild(PScore);
                    PScore.InnerText = K12.Data.Decimal.GetString(0);


                    XmlElement Score = doc.CreateElement("Score");
                    Extension_innerCreate.AppendChild(Score);
                    Score.InnerText = K12.Data.Decimal.GetString(score + System.Decimal.Parse(Extension_inner.SelectSingleNode("PScore").InnerText));


                    element.SelectSingleNode("Score").InnerText = Extension_inner.SelectSingleNode("Score").InnerText;


                }

            }

            target.Load(element);
        }

        //public static decimal? GetScore(this SCETakeRecord target)
        //{
        //    XmlElement element = target.ToXML();
        //    return K12.Data.Decimal.ParseAllowNull(element.SelectSingleNode("Score").InnerText);
        //}

        public static decimal? GetScore(this SCETakeRecord target)
        {
            XmlElement element = target.ToXML();

            XmlElement Extension = element.SelectSingleNode("Extension") as XmlElement;

            XmlElement Extension_inner = Extension.SelectSingleNode("Extension") as XmlElement;

            if (Extension_inner != null)
            {
                XmlElement eleCScore = Extension_inner.SelectSingleNode("CScore") as XmlElement;

                if (eleCScore != null)
                {
                    return K12.Data.Decimal.ParseAllowNull(eleCScore.InnerText);
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
            //return K12.Data.Decimal.ParseAllowNull(element.SelectSingleNode("Score").InnerText);
        }

    }
}
