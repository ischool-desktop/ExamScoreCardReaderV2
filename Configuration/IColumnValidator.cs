﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExamScoreCardReaderV2
{
    interface IColumnValidator
    {
        bool IsValid(string input);
        string GetErrorMessage();
    }
}
