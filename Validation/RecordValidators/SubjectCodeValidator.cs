﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExamScoreCardReaderV2.Model;
using ExamScoreCardReaderV2.Mapper;

namespace ExamScoreCardReaderV2.Validation.RecordValidators
{
    internal class SubjectCodeValidator : IRecordValidator<RawData>
    {
        #region IRecordValidator<RawData> 成員
        public string Validate(RawData record)
        {
            if (SubjectCodeMapper.Instance.CheckCodeExists(record.SubjectCode))
                return string.Empty;
            else
                return string.Format("科目代碼「{0}」不存在。", record.SubjectCode);
        }
        #endregion
    }
}
