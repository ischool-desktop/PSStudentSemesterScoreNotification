using Aspose.Words;
using FISCA.Data;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using JHSchool.Data;
using JHSchool.Evaluation.Mapping;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace PSStudentSemesterScoreNotification
{
	public partial class MainForm : BaseForm
	{
		//主文件
		private Document _doc;

		//單頁範本
		private Document _template;

		//進入類型判斷( 學生 or 班級 )
		private EnterType _enterType;

		//字串類型 學年度、學期
		private string _schoolYear, _semester;

		private int schoolYear, semester;

		//等第對照
		private DegreeMapper _degreeMapper;

		private QueryHelper queryHelper = new QueryHelper();

		BackgroundWorker BGW = new BackgroundWorker();
		private string CadreConfig = "PSStudentSemesterScoreNotification.cs";

		internal static void Run(EnterType enterType)
		{
			new MainForm(enterType).ShowDialog();
		}

		private void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			try
			{
				_doc.Sections.RemoveAt(0);
				_doc.MailMerge.RemoveEmptyParagraphs = true;
				_doc.MailMerge.DeleteFields(); //刪除未使用的功能變數
				btn_Print.Enabled = true;
			}
			catch
			{

			}

			try
			{
				SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

				SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
				SaveFileDialog1.FileName = "學期學習表現通知單(國小)" + string.Format("{0:MM_dd_yy_H_mm_ss}", DateTime.Now);

				if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
				{
					_doc.Save(SaveFileDialog1.FileName);
					Process.Start(SaveFileDialog1.FileName);
					MotherForm.SetStatusBarMessage("學期學習表現通知單(國小),列印完成!!");
				}
				else
				{
					FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
					return;
				}
			}
			catch
			{
				FISCA.Presentation.Controls.MsgBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
				MotherForm.SetStatusBarMessage("檔案儲存錯誤,請檢查檔案是否開啟中!!");
			}
		}

		internal List<JHStudentRecord> GetStudents()
		{
			if (_enterType == EnterType.Student)
				return JHStudent.SelectByIDs(K12.Presentation.NLDPanels.Student.SelectedSource);
			else
			{
				List<JHStudentRecord> list = new List<JHStudentRecord>();
				// 取得班級學生(一般和輟學)                
				foreach (JHClassRecord cla in JHClass.SelectByIDs(K12.Presentation.NLDPanels.Class.SelectedSource))
				{
					foreach (JHStudentRecord stud in cla.Students)
						if (stud.Status == StudentRecord.StudentStatus.一般 || stud.Status == StudentRecord.StudentStatus.輟學)
							list.Add(stud);
				}
				return list;
			}
		}

		public MainForm(EnterType enterType)
		{
			InitializeComponent();
			_enterType = enterType;
			InitializeSemester();
			_degreeMapper = new DegreeMapper();
			BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
			BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
		}

		private void InitializeSemester()
		{
			try
			{
				schoolYear = int.Parse(School.DefaultSchoolYear);
				semester = int.Parse(School.DefaultSemester);
				for (int i = -3; i <= 2; i++)
					cboSchoolYear.Items.Add(schoolYear + i);
				cboSemester.Items.Add(1);
				cboSemester.Items.Add(2);

				cboSchoolYear.Text = schoolYear.ToString();
				cboSemester.Text = semester.ToString();
			}
			catch (Exception ex)
			{
				MsgBox.Show("必須選擇為數字");
			}
		}

		private string ServerTime()
		{
			QueryHelper Sql = new QueryHelper();
			DataTable dtable = Sql.Select("select now()"); //取得時間
			DateTime dt = DateTime.Now;
			DateTime.TryParse("" + dtable.Rows[0][0], out dt); //Parse資料
			string ComputerSendTime = dt.ToString("yyyy/MM/dd"); //最後時間

			return ComputerSendTime;
		}

		private void btn_Print_Click(object sender, EventArgs e)
		{
			_schoolYear = cboSchoolYear.Text;
			_semester = cboSemester.Text;
			BGW.RunWorkerAsync();
		}

		private void BGW_DoWork(object sender, DoWorkEventArgs e)
		{
			_doc = new Document();
			#region 資料建立
			string DatePrint = ServerTime();
			string schoolName = K12.Data.School.ChineseName;
			List<JHStudentRecord> students = GetStudents();
			List<SemesterScoreRecord> ssrl = K12.Data.SemesterScore.SelectByStudents(students);
			List<SemesterHistoryRecord> semesterHistory = SemesterHistory.SelectByStudents(students);
			Dictionary<string, SemesterHistoryRecord> studentHistoryDic = new Dictionary<string, SemesterHistoryRecord>();

			foreach (SemesterHistoryRecord shr in semesterHistory)
			{
				if (!studentHistoryDic.ContainsKey(shr.RefStudentID))
				{
					studentHistoryDic.Add(shr.RefStudentID, shr);
				}
			}

			List<JHSemesterScoreRecord> jss = JHSchool.Data.JHSemesterScore.SelectByStudents(students);
			Dictionary<string, List<DomainScore>> studentDomainScore = new Dictionary<string, List<DomainScore>>();

			//老師評語
			Dictionary<string, JHMoralScoreRecord> dmsr = new Dictionary<string, JHMoralScoreRecord>();
			List<JHMoralScoreRecord> msrl = JHMoralScore.SelectByStudentIDs(students.Select(x => x.ID));
			foreach (JHMoralScoreRecord JHmsr in msrl)
			{
				if (!dmsr.ContainsKey(JHmsr.RefStudentID + "#" + JHmsr.SchoolYear + "#" + JHmsr.Semester))
					dmsr.Add(JHmsr.RefStudentID + "#" + JHmsr.SchoolYear + "#" + JHmsr.Semester, JHmsr);
			}

			foreach (JHSemesterScoreRecord ss in jss)
			{
				foreach (DomainScore ds in ss.Domains.Values)
				{
					string studentID = ds.RefStudentID;
					string domain = ds.Domain;
					string text = ds.Text;

					if (ds.SchoolYear + "" == _schoolYear && ds.Semester + "" == _semester)
					{
						if (!studentDomainScore.ContainsKey(studentID))
						{
							studentDomainScore.Add(studentID, new List<DomainScore>());
						}
						studentDomainScore[studentID].Add(ds);
					}
				}
			}


			//科目資料處理
			Dictionary<string, string> studentSubjectDic = new Dictionary<string, string>();
			foreach (SemesterScoreRecord ssr in ssrl)
			{
				if (ssr.Subjects.ContainsKey("國語"))
				{
					decimal? score = ssr.Subjects["國語"].Score;

					if (ssr.SchoolYear + "" == _schoolYear && ssr.Semester + "" == _semester)
					{
						if (!studentSubjectDic.ContainsKey(ssr.RefStudentID + "_國語") && score.HasValue)
							studentSubjectDic.Add(ssr.RefStudentID + "_國語", _degreeMapper.GetDegreeByScore(score.Value));
					}
				}

				if (ssr.Subjects.ContainsKey("英語"))
				{
					decimal? score = ssr.Subjects["英語"].Score.Value;

					if (ssr.SchoolYear + "" == _schoolYear && ssr.Semester + "" == _semester)
					{
						if (!studentSubjectDic.ContainsKey(ssr.RefStudentID + "_英語") && score.HasValue)
							studentSubjectDic.Add(ssr.RefStudentID + "_英語", _degreeMapper.GetDegreeByScore(score.Value));
					}
				}

				if (ssr.Subjects.ContainsKey("本土語"))
				{
					decimal? score = ssr.Subjects["本土語"].Score.Value;

					if (ssr.SchoolYear + "" == _schoolYear && ssr.Semester + "" == _semester)
					{
						if (!studentSubjectDic.ContainsKey(ssr.RefStudentID + "_本土語") && score.HasValue)
							studentSubjectDic.Add(ssr.RefStudentID + "_本土語", _degreeMapper.GetDegreeByScore(score.Value));
					}
				}
			}
			#endregion

			#region 列印資料

			Campus.Report.ReportConfiguration ConfigurationInCadre = new Campus.Report.ReportConfiguration(CadreConfig);
			if (ConfigurationInCadre.Template == null)
			{
				//如果範本為空,則建立一個預設範本
				ConfigurationInCadre.Template = new Campus.Report.ReportTemplate(Properties.Resources.學期學習表現通知單_國小_, Campus.Report.TemplateType.Word);
			}

			_template = new Document(new MemoryStream(ConfigurationInCadre.Template.ToBinary()));
			
			//new MemoryStream(ConfigurationInCadre.Template.ToBinary)
			foreach (JHStudentRecord JHsr in students)
			{
				string personalDays = "", sickDays = "";
				string sql = "select ref_student_id,personal_days,sick_days from $ischool.elementaryabsence where ref_student_id in (" + JHsr.ID + ") and school_year=" + _schoolYear + " and semester=" + _semester;
				DataTable dt = queryHelper.Select(sql);
				foreach (DataRow row in dt.Rows)
				{
					 personalDays = "" + row["personal_days"];
					 sickDays = "" + row["sick_days"];
				}
				
				string teacherRemark = "";
				Document perPage = _template.Clone();
				Dictionary<String, String> mergeDic = new Dictionary<string, string>();

				mergeDic.Add("列印日期", DatePrint);
				mergeDic.Add("學校名稱", schoolName);
				mergeDic.Add("學年度", _schoolYear);
				mergeDic.Add("學期", _semester);
				mergeDic.Add("班級", JHsr.Class.Name);
				mergeDic.Add("座號", JHsr.SeatNo + "");
				mergeDic.Add("姓名", JHsr.Name);
				mergeDic.Add("事假日數",personalDays);
				mergeDic.Add("病假日數", sickDays);

				/// <summary>
				/// 列印學生出席日數
				/// </summary>
				foreach (SemesterHistoryItem shi in studentHistoryDic[JHsr.ID].SemesterHistoryItems)
				{
					if (shi.SchoolYear + "" == _schoolYear && shi.Semester + "" == _semester)
						mergeDic.Add("應出席日數", shi.SchoolDayCount + "日");
				}

				string 國語 = studentSubjectDic.ContainsKey(JHsr.ID + "_國語") ? studentSubjectDic[JHsr.ID + "_國語"] : "";
				mergeDic.Add("語文_國語_等第", 國語);
				string 英語 = studentSubjectDic.ContainsKey(JHsr.ID + "_英語") ? studentSubjectDic[JHsr.ID + "_英語"] : "";
				mergeDic.Add("語文_英語_等第", 英語);
				string 本土語 = studentSubjectDic.ContainsKey(JHsr.ID + "_本土語") ? studentSubjectDic[JHsr.ID + "_本土語"] : "";
				mergeDic.Add("語文_本土語_等第", 本土語);

				if (studentDomainScore.ContainsKey(JHsr.ID))
				{
					foreach (DomainScore ds in studentDomainScore[JHsr.ID])
					{
						if (ds.Score.HasValue)
						{
							mergeDic.Add(ds.Domain + "_等第", _degreeMapper.GetDegreeByScore(ds.Score.Value));
						}
						mergeDic.Add(ds.Domain + "_文字", ds.Text);
					}
				}

				/// <summary>
				/// 列印老師評語
				/// </summary>
				//List<AttendanceRecord> studentAttend = Attendance.SelectBySchoolYearAndSemester(students, schoolYear, semester);
				//Dictionary<string, string> studentAttendDic = new Dictionary<string, string>();
				//List<JHMoralScoreRecord> Jmsr = JHMoralScore.SelectBySchoolYearAndSemester(students, schoolYear, semester);
				

				if (dmsr.ContainsKey(JHsr.ID + "#" + _schoolYear + "#" + _semester))
				{
					JHMoralScoreRecord msr = dmsr[JHsr.ID + "#" + _schoolYear + "#" + _semester];
					XmlElement Element;
					if (msr.TextScore != null)
					{
						Element = msr.TextScore.SelectSingleNode("DailyLifeRecommend") as XmlElement;
						if (Element != null)
							teacherRemark = Element.GetAttribute("Description");
					}
					mergeDic.Add("老師評語", teacherRemark);
				}

				perPage.MailMerge.CleanupOptions = Aspose.Words.Reporting.MailMergeCleanupOptions.RemoveEmptyParagraphs;
				perPage.MailMerge.Execute(mergeDic.Keys.ToArray<string>(), mergeDic.Values.ToArray<object>());
				perPage.MailMerge.DeleteFields();

				_doc.Sections.Add(_doc.ImportNode(perPage.FirstSection, true));
			}
			#endregion
		}

		private void btn_Exit_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			//取得設定檔
			Campus.Report.ReportConfiguration ConfigurationInCadre = new Campus.Report.ReportConfiguration(CadreConfig);
			//畫面內容(範本內容,預設樣式
			Campus.Report.TemplateSettingForm TemplateForm;
			if (ConfigurationInCadre.Template != null)
			{
				TemplateForm = new Campus.Report.TemplateSettingForm(ConfigurationInCadre.Template, new Campus.Report.ReportTemplate(Properties.Resources.學期學習表現通知單_國小_, Campus.Report.TemplateType.Word));
			}
			else
			{
				ConfigurationInCadre.Template = new Campus.Report.ReportTemplate(Properties.Resources.學期學習表現通知單_國小_, Campus.Report.TemplateType.Word);
				TemplateForm = new Campus.Report.TemplateSettingForm(ConfigurationInCadre.Template, new Campus.Report.ReportTemplate(Properties.Resources.學期學習表現通知單_國小_, Campus.Report.TemplateType.Word));
			}

			//預設名稱
			TemplateForm.DefaultFileName = "學期學習表現通知單(國小)樣板";
			//如果回傳為OK
			if (TemplateForm.ShowDialog() == DialogResult.OK)
			{
				//設定後樣試,回傳
				ConfigurationInCadre.Template = TemplateForm.Template;
				//儲存
				ConfigurationInCadre.Save();
			}
		}
	}

	public enum EnterType
	{
		Student, Class
	}
}
