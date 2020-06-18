using System;
using System.Windows.Forms;
using System.Configuration;
using System.Xml;
using System.IO;

namespace TableGenerater
{
    public partial class TableGenerater : Form
    {
        public TableGenerater()
        {
            InitializeComponent();

            string projectPathList = ConfigurationManager.AppSettings["ProjectPathList"];
            if (false == string.IsNullOrEmpty(projectPathList))
            {
                string[] paths = projectPathList.Split(',');

                this.comboBoxProject.Items.AddRange(paths);
                this.comboBoxProject.SelectedIndex = 0;
            }

            string excelPathList = ConfigurationManager.AppSettings["ExcelPathList"];
            if (false == string.IsNullOrEmpty(excelPathList))
            {
                string[] paths = excelPathList.Split(',');

                this.comboBoxExcel.Items.AddRange(paths);
                this.comboBoxExcel.SelectedIndex = 0;
            }
        }

        private void buttonFindProject_Click(object sender, EventArgs e)
        {
            LogToTextBox("클라이언트 프로젝트의 루트 폴더를 선택하세요.");

            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();

            string selected = dialog.SelectedPath;

            this.comboBoxProject.Items.Insert(0, selected);
            this.comboBoxProject.SelectedIndex = 0;
        }

        private void buttonFindExcel_Click(object sender, EventArgs e)
        {
            LogToTextBox("엑셀 데이터가 존재하는 폴더를 선택하세요.");

            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();

            string selected = dialog.SelectedPath;

            this.comboBoxExcel.Items.Insert(0, selected);
            this.comboBoxExcel.SelectedIndex = 0;
        }

        private void buttonGenerater_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(this.comboBoxProject.SelectedItem as string) || string.IsNullOrEmpty(this.comboBoxExcel.SelectedItem as string))
            {
                LogToTextBox("클라이언트 프로젝트 및 데이터 경로 확인이 필요합니다.");
                return;
            }

            LogToTextBox("생성 시작!!");

            string codePath = this.comboBoxProject.SelectedItem as string;
            codePath += @"\Assets\Scripts\TableHeader";
            CheckDirectory(codePath);

            string binaryPath = this.comboBoxProject.SelectedItem as string;
            binaryPath += @"\Assets\Resources\Text";
            CheckDirectory(binaryPath);

            Generater.Headergenerater generater = new Generater.Headergenerater();

            generater.Initialize(this.comboBoxExcel.SelectedItem as string);
            foreach (var iter in generater.GenerateCode(codePath))
            {
                LogToTextBox(iter);
            }

            foreach (var iter in generater.GenerateBinary(binaryPath))
            {
                LogToTextBox(iter);
            }

            if (true == this.checkBoxStreamingAssets.Checked)
            {
                var streamingAssetsPath = this.comboBoxProject.SelectedItem as string;
                streamingAssetsPath += @"\Assets\StreamingAssets";

                foreach (var iter in generater.GenerateBinary(streamingAssetsPath))
                {
                    LogToTextBox(iter);
                }
            }

            SaveToXml(this.comboBoxProject, "ProjectPathList");
            SaveToXml(this.comboBoxExcel, "ExcelPathList");

            LogToTextBox("생성 성공!! " + DateTime.Now.ToString());
        }

        /// <summary>
        /// 콤보 박스 정보를 xml 문서에 저장한다.
        /// </summary>
        /// <param name="combox">경로 목록이 저장된 콤보 박스</param>
        /// <param name="keyValue">경로들이 저장될 xml 노드</param>
        private void SaveToXml(ComboBox combox, string keyValue)
        {
            string paths = string.Empty;

            foreach (var item in combox.Items)
            {
                if (false == string.IsNullOrEmpty(paths))
                {
                    paths += ",";
                }

                paths += (string)item;
            }

            string appConfigPath = Application.StartupPath + @"\TableGenerater.exe.config";

            #region 사용자 설정이 존재하지 않다면 생성한다.
            if (false == File.Exists(appConfigPath))
            {
                XmlDocument doc = new XmlDocument();

                XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
                XmlElement root = doc.DocumentElement;
                doc.InsertBefore(xmlDeclaration, root);

                XmlElement configuration = doc.CreateElement(string.Empty, "configuration", string.Empty);
                doc.AppendChild(configuration);

                XmlElement startup = doc.CreateElement(string.Empty, "startup", string.Empty);
                configuration.AppendChild(startup);

                XmlElement supportedRuntime = doc.CreateElement("supportedRuntime");
                supportedRuntime.SetAttribute("version", "v4.0");
                supportedRuntime.SetAttribute("sku", ".NETFramework,Version=v4.5");
                startup.AppendChild(supportedRuntime);

                XmlElement appSettings = doc.CreateElement(string.Empty, "appSettings", string.Empty);
                configuration.AppendChild(appSettings);

                var add_1 = doc.CreateElement("add");
                add_1.SetAttribute("key", "ProjectPathList");
                add_1.SetAttribute("value", string.Empty);

                appSettings.AppendChild(add_1);

                var add_2 = doc.CreateElement("add");
                add_2.SetAttribute("key", "ExcelPathList");
                add_2.SetAttribute("value", string.Empty);

                appSettings.AppendChild(add_2);

                doc.Save(appConfigPath);
            }
            #endregion

            XmlDocument xml = new XmlDocument();
            xml.Load(appConfigPath);

            XmlNode node = xml.SelectSingleNode("configuration/appSettings");

            for (var i = 0; i < node.ChildNodes.Count; ++i)
            {
                if (true == string.Equals(keyValue, node.ChildNodes[i].Attributes["key"].Value, StringComparison.OrdinalIgnoreCase))
                {
                    node.ChildNodes[i].Attributes["value"].Value = paths;
                }
            }

            xml.Save(appConfigPath);
        }

        /// <summary>
        /// 텍스트 박스를 갱신하고 업데이트한다.
        /// </summary>
        /// <param name="message">갱신 메시지</param>
        private void LogToTextBox(string message)
        {
            this.textBox.Text = message;
            this.textBox.Update();
        }

        /// <summary>
        /// 디렉토리 존재 여부를 확인하고 필요 시 생성한다.
        /// </summary>
        /// <param name="path">디렉토리 경로</param>
        private void CheckDirectory(string path)
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(path);
            if (false == directoryInfo.Exists)
            {
                directoryInfo.Create();
            }
        }
    }
}
