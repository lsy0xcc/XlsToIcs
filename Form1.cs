using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;
using System.Text.RegularExpressions;
using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using static System.Console;

namespace XlsToIcs
{
    public partial class Form1 : Form
    {
        const String NotMondayMessage = "开学日期不是星期一";
        const String FileErrorMessage = "文件错误";
        String xlsFilePathString;
        
        public Form1()
        {
            InitializeComponent();
        }

        private void XlsFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = " Excel文件|*.xls";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                xlsFilePathString= Path.GetFullPath(fileDialog.FileName);//将选中的文件的路径传递给TextBox
                XlsFilePath.Text = xlsFilePathString;
            }
        }

        private void ToIcsFileButton_Click(object sender, EventArgs e)
        {
            String icsFileName;//ics文件名
            var courses = new List<Course>();//课程列表
            var calendar = new Calendar();//日历相关
            var serializer = new CalendarSerializer();
            var serializedCalendar = serializer.SerializeToString(calendar);
            var termStrat = new DateTime();//学期开始日
            string icsPath; //ics文件完整目录
            try
            {
                //得到ics文件名并且获得课程列表
                icsFileName = XlsToCourse(courses, XlsFilePath.Text);

                //判断开始日是否是星期一并生成学期开始日期
                if(dateTimePicker1.Value.DayOfWeek != DayOfWeek.Monday)
                {
                    throw new Exception(NotMondayMessage);
                }
                termStrat = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day, 0, 0, 0);
                CourseToIcs(courses, calendar,termStrat);
                serializer = new CalendarSerializer();


                serializedCalendar = serializer.SerializeToString(calendar);
                icsPath = Path.Combine(Path.GetDirectoryName(xlsFilePathString), icsFileName + ".ics");
                File.WriteAllText(icsPath, serializedCalendar);

                MessageBox.Show("成功生成文件");
            } catch(Exception ex)
            {
                if(ex.Message.CompareTo(NotMondayMessage) == 0)
                {
                    MessageBox.Show(NotMondayMessage);
                }
                else
                {
                    MessageBox.Show(FileErrorMessage);
                }
                
            }


        }

        static String XlsToCourse(List<Course> courses, String xlsFilePath)
        {

            String sheetName = "学生课表";
            int titleRow = 0;//标题的行
            int titleCell = 0;//标题的列
            int beginRowIndex = 2;//课表开始的行
            int beginCellIndex = 2;//课表开始的列
            int classNum = 6;//一天几堂课
            int daysOfWeek = 7; //一周有几天
            Regex rgx = new Regex(@"(<br\/>(?=考试时间))|(?<=周)(<\/br>)");//处理多余换行符的正则表达式
            Regex rgx2 = new Regex(@"</br>");//sb编译器不能按字符串拆分,把回车符换成特殊字符
            

            String temp = null;
            String[] array = null;//临时变量
            
            FileStream fs = new FileStream(xlsFilePath, FileMode.Open, FileAccess.Read);
            HSSFWorkbook wk = new HSSFWorkbook(fs);
            ISheet hs = wk.GetSheet(sheetName);
            for (int i = beginRowIndex; i < beginRowIndex + classNum; i++)
            {
                IRow ir = hs.GetRow(i);
                for (int j = beginCellIndex; j < beginCellIndex + daysOfWeek; j++)
                {
                    //删除多余的换行符
                    temp = rgx.Replace(ir.GetCell(j).ToString(), "");
                    //注意这里
                    temp = rgx2.Replace(temp, "$");
                    array = temp.Split('$');
                    //判断是否输入正确
                    if (array.Length % 2 == 0)
                    {
                        for (int k = 0; k < array.Length - 1; k += 2)
                        {
                            courses.Add(new Course(i + 1 - beginRowIndex, j + 1 - beginCellIndex, array[k], array[k + 1]));
                        }
                    }
                    else
                    {
                        // dosomething
                    }

                }
            }

            String icsFileName = hs.GetRow(titleRow).GetCell(titleCell).ToString();
            fs.Close();
            return icsFileName;
        }

        static void CourseToIcs(List<Course> courses, Calendar calendar, DateTime termStart)
        {
            //DateTime termStart = new DateTime(2019, 2, 25, 0, 0, 0);
            DateTime beginDate = termStart.AddDays(-8);
            

            int[,] CourseTime = new int[6, 4] { { 8, 0, 9, 45 }, { 10, 00, 11, 45 }, { 13, 45, 15, 30 },
                { 15, 45, 17, 30 }, { 18, 30, 20, 15 }, { 20, 30, 22, 15 } };
            int[,] ExperTime = new int[4, 4] { { 7, 20, 9, 50 }, { 10, 00, 12, 30 }, { 13, 00, 15, 30 },
                { 15, 40, 18, 30 }};

            DateTime begin;
            DateTime end;
            CalendarEvent e;
            foreach (Course c in courses)
            {
                foreach (int weeknum in c.Weeks)
                {
                    begin = beginDate.AddDays(weeknum * 7 + c.DayOfWeek);
                    end = beginDate.AddDays(weeknum * 7 + c.DayOfWeek);
                    if (c.courseType == Course.Type.course)
                    {
                        begin = beginDate.AddDays(weeknum * 7 + c.DayOfWeek)
                            .AddHours(CourseTime[c.CourseNum - 1, 0]).AddMinutes(CourseTime[c.CourseNum - 1, 1]);
                        end = beginDate.AddDays(weeknum * 7 + c.DayOfWeek)
                            .AddHours(CourseTime[c.CourseNum - 1, 2]).AddMinutes(CourseTime[c.CourseNum - 1, 3]);

                    }
                    else if (c.courseType == Course.Type.experiment)
                    {
                        begin = beginDate.AddDays(weeknum * 7 + c.DayOfWeek)
                            .AddHours(ExperTime[c.CourseNum - 1, 0]).AddMinutes(ExperTime[c.CourseNum - 1, 1]);
                        end = beginDate.AddDays(weeknum * 7 + c.DayOfWeek)
                            .AddHours(ExperTime[c.CourseNum - 1, 2]).AddMinutes(ExperTime[c.CourseNum - 1, 3]);
                    }
                    e = new CalendarEvent
                    {
                        Start = new CalDateTime(begin),
                        End = new CalDateTime(end),
                        Summary = $"第{c.CourseNum}节@{c.Position} {c.CourseName}",
                    };
                    calendar.Events.Add(e);


                }
            }

        }

    }

    class Course
    {
        public enum Type { course, experiment, examination }

        public String CourseName { get; set; }//课程名
        public String CourseInfo { get; set; }//其他
        public int DayOfWeek { get; set; }//星期
        public int CourseNum { get; set; }//课程节数

        public List<int> Weeks = new List<int>();//上课周数
        public String ExamTime;//考试时间
        public String Position;//上课地点
        public Type courseType;//课程种类

        public Course(int courseNum, int dayOfWeek, String courseName, String courseInfo)
        {
            this.CourseInfo = courseInfo;
            this.CourseName = courseName;
            this.DayOfWeek = dayOfWeek;
            this.CourseNum = courseNum;
            SetOtherInfo();
        }

        public void SetOtherInfo()
        {
            var examRegex = new Regex(@"^\[考试\]");
            var experRegex = new Regex(@"\(实验\)$");

            var examTimeRegex = new Regex(@"(?<=考试时间:)(\d\d:\d\d-\d\d:\d\d)");
            var positionRegex = new Regex(@"(?<=周(?!，)).*");
            var weekRegex = new Regex(@"(?<=\[).*?(?=\])");
            var weekNumREgex = new Regex(@"((\d+)-)*(\d+)((单|双)*)");

            #region 设置课程属性 如果是考试改变考试时间
            if (examRegex.IsMatch(CourseName))
            {
                this.courseType = Type.examination;
                foreach (Match m in examTimeRegex.Matches(CourseInfo))
                {
                    ExamTime = m.Value;
                }

            }
            else if (experRegex.IsMatch(CourseName))
            {
                this.courseType = Type.experiment;
            }
            else
            {
                this.courseType = Type.course;
            }
            #endregion

            //上课地点
            Position = positionRegex.Match(CourseInfo).Value;

            #region 上课周数
            String[] temp;
            String tempString;
            bool odd = true;
            bool even = true;
            Match match;
            foreach (Match m in weekRegex.Matches(CourseInfo))
            {
                temp = m.ToString().Split('，');
                foreach (String s in temp)
                {
                    odd = true;
                    even = true;
                    try
                    {
                        match = weekNumREgex.Match(s);
                        tempString = weekNumREgex.Match(s).Value;
                        if (match.Groups[4].Value.CompareTo("单") == 0)
                        {
                            even = false;
                        }
                        else if (match.Groups[4].Value.CompareTo("双") == 0)
                        {
                            odd = false;
                        }

                        if (match.Groups[2].Value.CompareTo("") == 0)
                        {
                            Weeks.Add(int.Parse(match.Groups[3].Value));
                        }
                        else
                        {
                            for (int i = int.Parse(match.Groups[2].Value); i < int.Parse(match.Groups[3].Value); i++)
                            {
                                if ((i % 2 == 0 && even) || (i % 2 == 1 && odd))
                                {
                                    Weeks.Add(i);
                                }
                            }
                        }


                    }
                    catch (Exception e)
                    {
                        WriteLine(e);
                    }
                }

            }
            #endregion


        }
    }
}
