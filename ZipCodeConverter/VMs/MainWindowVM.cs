using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Windows.Input;
using System.Windows;
using System.Runtime.Serialization.Json;
using Microsoft.Win32;
using System.IO;
using Ionic.Zip;
using Microsoft.VisualBasic.FileIO;
using System.Net;

namespace ZipCodeConverter.VMs
{
    public class MainWindowVM : INotifyPropertyChanged
    {
        public MainWindowVM()
        {
            SaveDirectory = @"c:\json\";
            DownloadURL = "http://www.post.japanpost.jp/zipcode/dl/kogaki/zip/ken_all.zip";
            IsDownload = false;
            ProgressVisibility = Visibility.Collapsed;
        }

        #region Properties
        
        string saveDirectory;
        public string SaveDirectory
        {
            get { return saveDirectory; }
            set
            {
                saveDirectory = value;
                FirePropertyChanged("SaveDirectory");
            }
        }

        string downloadURL;
        public string DownloadURL
        {
            get { return downloadURL; }
            set
            {
                downloadURL = value;
                FirePropertyChanged("DownloadURL");
            }
        }

        bool isDownload;
        /// <summary>
        /// WebからZIPファイルをダウンロードし実行するかローカルファイルに対して実行するか
        /// </summary>
        public bool IsDownload
        {
            get { return isDownload; }
            set
            {
                isDownload = value;
                FirePropertyChanged("IsDownload");
            }
        }

        public CommandBase ButtonCommand
        {
            get 
            {
                return new CommandBase(x => 
                {
                    if (IsDownload)
                    {
                        DownloadFromWeb();
                    }
                    else
                    {
                        LoadLocalFile();
                    }
                },
                y => { return true; });
            }
        }

        #region progressBar系プロパティ
        int progressValue;
        public int ProgressValue
        {
            get { return progressValue; }
            set
            {
                progressValue = value;
                FirePropertyChanged("ProgressValue");
            }
        }

        int progressMaxNum;
        public int ProgressMaxNum
        {
            get { return progressMaxNum; }
            set
            {
                progressMaxNum = value;
                FirePropertyChanged("ProgressMaxNum");
            }
        }

        string progressStatus;
        public string ProgressStatus
        {
            get { return progressStatus; }
            set
            {
                progressStatus = value;
                FirePropertyChanged("ProgressStatus");
            }
        }

        Visibility progressVisibility;
        public Visibility ProgressVisibility
        {
            get { return progressVisibility; }
            set
            {
                progressVisibility = value;
                FirePropertyChanged("ProgressVisibility");
                FirePropertyChanged("MaskColor");
            }
        }

        #endregion
        
        /// <summary>
        /// 処理実行中のマスク
        /// 別階層にしてIsEnable=falseでも
        /// </summary>
        public string MaskColor
        {
            get
            {
                if (ProgressVisibility == Visibility.Visible)
                {
                    return "#8000";
                }
                else
                {
                    return null;
                }
            }
        }

        #endregion

        #region 処理
        void DownloadFromWeb()
        {
            ProgressMaxNum = 10;
            ProgressValue = 0;
            ProgressVisibility = Visibility.Visible;
            ProgressStatus = "ファイルダウンロード中";
            WebClient client = new WebClient();
            client.DownloadDataCompleted += OnDownloadCompleted;
            client.DownloadDataAsync(new Uri(DownloadURL));
        }

        void OnDownloadCompleted(object sender, DownloadDataCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show("ダウンロードに失敗しました : " + e.Error.Message);
                return;
            }
            ExecWorker(new MemoryStream(e.Result));
        }

        /// <summary>
        /// ローカルからZipファイルを読み込み処理を実行
        /// </summary>
        void LoadLocalFile()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            
            dialog.FileName = "";
            dialog.DefaultExt = "*.zip";
            if (dialog.ShowDialog() == true)
            {
                ProgressStatus = "ファイル展開中";
                ProgressMaxNum = 10;
                ProgressValue = 0;
                ProgressVisibility = Visibility.Visible;
                ExecWorker(new FileStream(dialog.FileName, FileMode.Open));
            }
            else
            {
                ProgressVisibility = System.Windows.Visibility.Collapsed;
            }
        }

        /// <summary>
        /// ZipFileStream内の先頭ファイルのStreamを返す
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        MemoryStream ExtractToStream(Stream stream)
        {
            using (ZipFile zip = ZipFile.Read(stream))
            {
                var exStream = new MemoryStream();
                zip.First().Extract(exStream);//無いと死ぬ
                exStream.Position = 0;
                return exStream;
            }
        }

        List<Models.Address> addresses = new List<Models.Address>();

        void ConvertToJSON(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            try
            {
                using (Stream zipStream = e.Argument as Stream)
                {
                    using (MemoryStream exStream = ExtractToStream(zipStream))
                    {
                        string directory = addDelimiter(SaveDirectory) + @"tmp\";

                        addresses.Clear();
                        TextFieldParser parser = new TextFieldParser(exStream, Encoding.GetEncoding("Shift_JIS"));
                        parser.TextFieldType = FieldType.Delimited;
                        parser.SetDelimiters(",");

                        ProgressStatus = "パース中";

                        ProgressMaxNum = 100;
                        ProgressValue = 0;

                        int count = 0;
                        //CSVファイルを1行ずつパースしてcollectionに追加
                        while (parser.EndOfData == false)
                        {
                            string[] columns = parser.ReadFields();
                            addresses.Add(new Models.Address(columns));

                            count++;

                            //1件ごとにやると細かすぎてうまくいかないので
                            if ((count % 1233) == 0)
                            {
                                Worker.ReportProgress(1);
                            }
                        }

                        ProgressMaxNum = addresses.GroupBy(x => x.ZipCode.Substring(0, 3)).Count();
                        ProgressValue = 0;

                        ProgressStatus = "個別ファイル作成中";

                        //フォルダがなければ作成
                        if (!System.IO.Directory.Exists(directory))
                        {
                            System.IO.Directory.CreateDirectory(directory);
                        }

                        //郵便番号の上3ケタごとに出力
                        foreach (var row in addresses.GroupBy(x => x.ZipCode.Substring(0, 3)))
                        {
                            //重複は除外しファイル出力
                            SerializeToJson(row.Distinct(new AddressComparer()).OrderBy(x => x.ZipCode).ToList(), directory + "zip" + row.Key + ".txt");
                            Worker.ReportProgress(1);
                        }

                        ProgressStatus = "ZIP圧縮中";

                        //ZipFileを作成する
                        using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
                        {
                            zip.AddFiles(System.IO.Directory.GetFiles(directory), "");
                            zip.Save(addDelimiter(SaveDirectory) + @"zipcode.zip");
                            System.IO.Directory.Delete(directory, true);
                        }
                    }
                }
                MessageBox.Show("完了");
            }
            catch (Exception ex)
            {
                MessageBox.Show("失敗 : " + ex.Message);
            }
            finally
            {
                ProgressVisibility = System.Windows.Visibility.Collapsed;
            }
        }

        string addDelimiter(string path)
        {
            if (path.EndsWith(@"\")) return path;
            return path + @"\";
        }

        /// <summary>
        /// 受け取ったクラスのインスタンスをJson形式でファイルに保存
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="val"></param>
        /// <param name="fileName"></param>
        public static void SerializeToJson<T>(T val, string fileName) where T : class
        {
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
            using (var stream = new FileStream(fileName, FileMode.Create))
            {
                serializer.WriteObject(stream, val);
                stream.Flush();
            }
        }

        /// <summary>
        /// 郵便番号重複チェック
        /// 郵便番号・住所1、住所2が同じなら重複として弾く
        /// </summary>
        class AddressComparer : EqualityComparer<Models.Address>
        {
            public override bool Equals(Models.Address a1, Models.Address a2)
            {
                if (Object.Equals(a1, a2))
                {
                    return true;
                }

                if (a1 == null || a2 == null)
                {
                    return false;
                }

                return (a1.ZipCode == a2.ZipCode
                    && a1.AddressItems[0] == a2.AddressItems[0]
                    && a1.AddressItems[1] == a2.AddressItems[1]
                    );
            }

            public override int GetHashCode(Models.Address a)
            {
                return a.ZipCode.GetHashCode();
            }
        }

        #region Worker系
        
        System.ComponentModel.BackgroundWorker Worker;

        // 現在値の更新
        void Worker_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            // プログレスバーの現在値を更新する
            ProgressValue++;
        }

        void ExecWorker(Stream stream)
        {
            Worker = new System.ComponentModel.BackgroundWorker();

            // イベントの登録
            Worker.DoWork += new System.ComponentModel.DoWorkEventHandler(ConvertToJSON);
            Worker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(Worker_ProgressChanged);

            // 進捗状況の報告をできるようにする
            Worker.WorkerReportsProgress = true;

            // バックグラウンド処理の実行
            Worker.RunWorkerAsync(stream);
        }

        #endregion
        #endregion

        #region INotifyPropertyChanged Members
        public event PropertyChangedEventHandler PropertyChanged;

        public void FirePropertyChanged(string name)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }
        #endregion
    }

    public class CommandBase : ICommand
    {
        public event EventHandler CanExecuteChanged;

        Func<object, bool> canExecute;
        Action<object> executeAction;
        bool canExecuteCache;

        public CommandBase(Action<object> executeAction, Func<object, bool> canExecute)
        {
            this.executeAction = executeAction;
            this.canExecute = canExecute;
        }

        #region ICommand Members

        public bool CanExecute(object parameter)
        {
            bool tempCanExecute = canExecute(parameter);

            if (canExecuteCache != tempCanExecute)
            {
                canExecuteCache = tempCanExecute;
                if (CanExecuteChanged != null)
                {
                    CanExecuteChanged(this, new EventArgs());
                }
            }

            return canExecuteCache;
        }

        public void Execute(object parameter)
        {
            executeAction(parameter);
        }
        #endregion
    }
}
