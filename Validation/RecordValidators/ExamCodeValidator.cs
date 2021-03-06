﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExamScoreCardReaderV2.Mapper;
using ExamScoreCardReaderV2.Model;

namespace ExamScoreCardReaderV2.Validation.RecordValidators
{
    internal class ExamCodeValidator : IRecordValidator<RawData>
    {
        #region IRecordValidator<RawData> 成員
        public string Validate(RawData record)
        {
            if (ExamCodeMapper.Instance.CheckCodeExists(record.ExamCode))
                return string.Empty;
            else
                return string.Format("試別代碼「{0}」不存在。", record.ExamCode);
        }
        #endregion
    }
}
