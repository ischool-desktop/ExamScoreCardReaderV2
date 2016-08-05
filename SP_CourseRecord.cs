using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace ExamScoreCardReaderV2
{
    public class SP_CourseRecord : K12.Data.CourseRecord
    {


        public new XmlElement Extensions
        {
            get
            {
                return base.Extensions;
            }
            set
            {
                base.Extensions = value;
            }
        }
    }
}
