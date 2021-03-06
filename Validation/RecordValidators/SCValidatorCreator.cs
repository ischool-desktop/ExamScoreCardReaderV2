﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExamScoreCardReaderV2.Model;
using K12.Data;

namespace ExamScoreCardReaderV2.Validation.RecordValidators
{
    internal class SCValidatorCreator
    {

        private StudentNumberDictionary _studentDict;
        private Dictionary<string, CourseRecord> _courseDict;
        private StudentCourseInfo _studentCourseInfo;
        private List<string> _attendCourseIDs;

        public StudentCourseInfo StudentCourseInfo { get { return _studentCourseInfo; } }
        public StudentNumberDictionary StudentNumberDictionary { get { return _studentDict; } }

        public List<string> AttendCourseIDs { get { return _attendCourseIDs; } }

        public SCValidatorCreator(List<StudentRecord> studentList, List<CourseRecord> courseList, List<SCAttendRecord> scaList)
        {
            _studentDict = new StudentNumberDictionary();
            _courseDict = new Dictionary<string, CourseRecord>();
            _studentCourseInfo = new StudentCourseInfo();
            _attendCourseIDs = new List<string>();

            foreach (StudentRecord student in studentList)
            {
                string studentNumber = student.StudentNumber;

                //studentNumber = GetStudentNumberFormat(studentNumber);

                if (!_studentDict.ContainsKey(studentNumber))
                    _studentDict.Add(studentNumber, student);

                _studentCourseInfo.Add(student);
            }

            foreach (CourseRecord course in courseList)
            {
                if (!_courseDict.ContainsKey(course.ID))
                    _courseDict.Add(course.ID, course);
            }

            //Linq
            var student_ids = from student in studentList select student.ID;

            //foreach (SCAttendRecord sc in SCAttend.Select(student_ids.ToList<string>(), null, null, "" + schoolYear, "" + semester))
            foreach (SCAttendRecord sc in scaList)
            {
                if (!_studentCourseInfo.ContainsID(sc.RefStudentID)) continue;
                if (!_courseDict.ContainsKey(sc.RefCourseID)) continue;

                if (!_attendCourseIDs.Contains(sc.RefCourseID))
                    _attendCourseIDs.Add(sc.RefCourseID);

                _studentCourseInfo.AddCourse(sc.RefStudentID, _courseDict[sc.RefCourseID]);
            }
        }

        //不明原因需要補 0，造成匯不進去，所以註解掉。
        //internal static string GetStudentNumberFormat(string studentNumber)
        //{
        //    #region 學號不足位，左邊補0
        //    int StudentNumberLength =Global.StudentNumberLenght;
        //    int s = StudentNumberLength - studentNumber.Length;
        //    if (s > 0)
        //        return studentNumber.PadLeft(StudentNumberLength, '0');
        //    else
        //        return studentNumber;
        //    #endregion
        //}

        internal IRecordValidator<DataRecord> CreateStudentValidator()
        {
            StudentValidator validator = new StudentValidator(_studentDict);
            return validator;
        }

        internal IRecordValidator<DataRecord> CreateSCAttendValidator()
        {
            SCAttendValidator validator = new SCAttendValidator(_studentCourseInfo);
            return validator;
        }
    }

    internal class StudentCourseInfo
    {
        private Dictionary<string, string> studentNumberTable; //StudentNumber -> ID
        private Dictionary<string, StudentRecord> studentTable; //ID -> Record
        private Dictionary<string, List<CourseRecord>> courseTable; //ID -> List of CourseRecord

        public StudentCourseInfo()
        {
            studentNumberTable = new Dictionary<string, string>();
            studentTable = new Dictionary<string, StudentRecord>();
            courseTable = new Dictionary<string, List<CourseRecord>>();
        }

        internal void Add(StudentRecord student)
        {
            if (string.IsNullOrEmpty(student.StudentNumber)) return;

            string studentNumber = student.StudentNumber;// SCValidatorCreator.GetStudentNumberFormat(student.StudentNumber);

            if (!studentNumberTable.ContainsKey(studentNumber))
                studentNumberTable.Add(studentNumber, student.ID);
            if (!studentTable.ContainsKey(student.ID))
                studentTable.Add(student.ID, student);
            if (!courseTable.ContainsKey(student.ID))
                courseTable.Add(student.ID, new List<CourseRecord>());
        }

        internal bool ContainsID(string id)
        {
            return studentTable.ContainsKey(id);
        }

        internal void AddCourse(string student_id, CourseRecord course)
        {
            if (!courseTable.ContainsKey(student_id)) return;
            courseTable[student_id].Add(course);
        }

        internal IEnumerable<CourseRecord> GetCourses(string sn)
        {
            if (!studentNumberTable.ContainsKey(sn)) return new List<CourseRecord>();
            string id = studentNumberTable[sn];
            if (!studentTable.ContainsKey(id)) return new List<CourseRecord>();

            return courseTable[id];
        }

        internal bool ContainsStudentNumber(string sn)
        {
            if (!studentNumberTable.ContainsKey(sn)) return false;
            else return true;
        }

        internal string GetStudentName(string sn)
        {
            if (!studentNumberTable.ContainsKey(sn)) return "<查無姓名>";
            string id = studentNumberTable[sn];
            if (!studentTable.ContainsKey(id)) return "<查無姓名>";

            return studentTable[id].Name;
        }
    }

    internal class StudentNumberDictionary : Dictionary<string, StudentRecord> { }
}
