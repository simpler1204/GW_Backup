using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using Renci.SshNet;
using Renci.SshNet.Sftp;
using ExcelDataReader;
using System.Collections.Generic;

namespace GW_Backup
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            InitializeDataGridView();
        }

        private void InitializeDataGridView()
        {
            dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add("Host", "IP");
            dataGridView1.Columns.Add("Username", "ID");
            dataGridView1.Columns.Add("Status", "작업상태");
            dataGridView1.Rows.Clear();
        }

        private void btnBackup_Click(object sender, EventArgs e)
        {
            // 엑셀 파일 선택
            string excelPath = "";
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files|*.xlsx;*.xls";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    excelPath = ofd.FileName;
                }
                else
                {
                    MessageBox.Show("엑셀 파일을 선택하세요.");
                    return;
                }
            }

            // 엑셀에서 SSH 정보 읽기
            var sshList = ReadSshInfoListFromExcel(excelPath);
            if (sshList.Count == 0)
            {
                MessageBox.Show("엑셀 파일에 SSH 정보가 없습니다.");
                return;
            }

            dataGridView1.Rows.Clear();

            // 바탕화면에 SSH_Backups라는 상위 폴더 생성
            string backupRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "SSH_Backups"
            );
            if (!Directory.Exists(backupRoot))
                Directory.CreateDirectory(backupRoot);

            // 각 SSH 정보에 대해 백업 수행
            foreach (var ssh in sshList)
            {
                string host = ssh.Host;
                string username = ssh.Username;
                string password = ssh.Password;
                string remoteDirectory = "/root";
                string localDirectory = Path.Combine(
                    backupRoot,
                    $"ssh_{host}" // 폴더명이 ssh_아이피 형식
                );

                string status = "";
                try
                {
                    // 3초 timeout 설정
                    var connectionInfo = new ConnectionInfo(
                        host,
                        22,
                        username,
                        new PasswordAuthenticationMethod(username, password)
                    )
                    {
                        Timeout = TimeSpan.FromSeconds(1)
                    };

                    using (var sftp = new SftpClient(connectionInfo))
                    {
                        sftp.Connect();
                        if (!Directory.Exists(localDirectory))
                            Directory.CreateDirectory(localDirectory);

                        DownloadDirectory(sftp, remoteDirectory, localDirectory);
                        sftp.Disconnect();
                    }
                    status = "성공";
                }
                catch (Exception ex)
                {
                    status = "실패: " + ex.Message;
                }

                dataGridView1.Rows.Add(host, username, status);
            }
        }

        private List<SshInfo> ReadSshInfoListFromExcel(string filePath)
        {
            var list = new List<SshInfo>();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet();
                var table = result.Tables[0];
                for (int i = 1; i < table.Rows.Count; i++) // 0번은 헤더
                {
                    var row = table.Rows[i];
                    if (row[0] == null || string.IsNullOrWhiteSpace(row[0].ToString()))
                        continue;
                    list.Add(new SshInfo
                    {
                        Host = row[0]?.ToString(),
                        Username = row[1]?.ToString(),
                        Password = row[2]?.ToString()
                    });
                }
            }
            return list;
        }

        private void DownloadDirectory(SftpClient client, string source, string destination)
        {
            var files = client.ListDirectory(source);
            foreach (var file in files)
            {
                if (file.Name == "." || file.Name == "..") continue;
                string destPath = Path.Combine(destination, file.Name);

                if (file.IsDirectory)
                {
                    if (!Directory.Exists(destPath))
                        Directory.CreateDirectory(destPath);
                    DownloadDirectory(client, file.FullName, destPath);
                }
                else if (file.IsRegularFile)
                {
                    using (var fs = new FileStream(destPath, FileMode.Create))
                    {
                        client.DownloadFile(file.FullName, fs);
                    }
                }
            }
        }

        private class SshInfo
        {
            public string Host { get; set; }
            public string Username { get; set; }
            public string Password { get; set; }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            btnBackup_Click(sender, e);
        }
    }
}