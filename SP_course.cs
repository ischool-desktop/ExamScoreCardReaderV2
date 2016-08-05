using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExamScoreCardReaderV2
{
    public class SP_course : K12.Data.Course
    {
        public new static List<SP_CourseRecord> SelectByIDs(IEnumerable<string> CourseIDs)
        {
            return SelectByIDs<SP_CourseRecord>(CourseIDs);
        }

        
    }
}
