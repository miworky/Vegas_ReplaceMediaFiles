/*
 * 
 */


// NuGet で MediaInfo.Native をインストール
//   VS2019 ツール -> NuGetパッケージマネージャ -> ソリューションの NuGet パッケージの管理 
//   MediaInfo を検索 -> 以下の 2 つをインストール
//     MediaInfo.Native
//     MediaInfo.Wrapper

// 参照の追加で「MediaInfo.Wrapper.dll」を追加する

// これで、MediaInfo が使えるようになる



// Exif
// JPGはBitmapクラスを使えば Exif にアクセスできるが、Bitmapでは HEIF を開けない
// そこで、 GroupDocs.Metadata を使う
// が、有償のようなので、やめる


using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using ScriptPortal.Vegas;
using System.Globalization;

using System.Text.RegularExpressions;

namespace vegastest1
{
    static class ExifTag
    {
        public const int ShootingDateAndTime = 0x9003;

    }

    public class SearchMedia
    {
        private List<Tuple<DateTime, string>> mediaPoolMedias;

        public SearchMedia(List<Tuple<DateTime, string>> mediaPoolImages)
        {
            this.mediaPoolMedias = mediaPoolImages;
        }

        // 日時を指定して画像のパスを取得する
        public string FindMedia(DateTime dateTime)
        {
            foreach (var media in mediaPoolMedias)
            {
                DateTime thisDateTime = media.Item1;
                if (thisDateTime.CompareTo(dateTime) == 0)
                {
                    // 日時が完全一致するファイルが見つかった
                    return media.Item2;
                }
            }

            // なかった
            return "";
        }

    }

    public class EntryPoint
    {
        public void FromVegas(Vegas vegas)
        {
            // トラックに張り付けてあるみてねからダウンロードした画像ファイルを、高画質の元画像ファイルに置き換える
            // みてねからダウンロードした画像ファイルには、Exif はない




            // ダイアログを開き出力するログのファイルパスをユーザーに選択させる
            string saveFilePath = GetFilePath(vegas.Project.FilePath, "ReplaceMediaFiles");
            if (saveFilePath.Length == 0)
            {
                return;
            }

            System.IO.StreamWriter writer = new System.IO.StreamWriter(saveFilePath, false, Encoding.GetEncoding("Shift_JIS"));

            // メディアプールから置換対象のメディアファイルの一覧を作成する
            List<Tuple<DateTime, string>> mediaPoolMedias = new List<Tuple<DateTime, string>>();

            MediaPool mediaPool = vegas.Project.MediaPool;
            foreach (Media media in mediaPool)
            {
                DateTime? dateTime = null;

                string path = media.FilePath; // メディアのファイルパス



                if (IsImage(path))
                {
//                    GetShootingDateTimeImageUsingMediaInfo(path, writer);
//                    continue;

                    // 撮影日時を取得
                    dateTime = GetShootingDateTimeImage(path);
                }
                else if (IsVideo(path))
                {
                    // 撮影日時を取得
                    dateTime = GetShootingDateTimeVideo(path, writer);
                }
                else
                {
                    // 画像でも動画でもなければ無視
                    continue;
                }

                if (!dateTime.HasValue)
                {
                    // 撮影日時を取得できなければ無視
                    continue;
                }

                // メディアプール内のファイルの撮影日時とパスをログに出力する
                writer.WriteLine(toString(dateTime, path));

                mediaPoolMedias.Add(Tuple.Create(dateTime.Value, path));
            }

            SearchMedia searchMedia = new SearchMedia(mediaPoolMedias);

            // タイムラインに配置されている画像をみて、メディアプールに同じ日時の画像があるか調べる
            // 同じ日時の画像があれば置き換える
            Dictionary<string, string> pathPair = new Dictionary<string, string>();

            foreach (Track track in vegas.Project.Tracks)
            {
                if (track.IsAudio())
                {
                    // オーディオトラックは無視
                    continue;
                }

                foreach (TrackEvent trackEvent in track.Events)
                {
                    if (trackEvent.Takes.Count != 1)
                    {
                        // 複数テイクあるイベントは無視
                        // （すでにMediaの置き換えが行われているので）
                        continue;
                    }

                    Take take = trackEvent.Takes[0];
                    {

                        Media media = take.Media;
                        string path = media.FilePath; // メディアのファイルパス
                        if (!IsImage(path) && !IsVideo(path))
                        {
                            // 画像でも動画でもなければ無視
                            continue;
                        }

                        // 撮影日時を取得する。ただし、ファイル名に撮影日時が含まれている必要がある
                        DateTime? dateTime = GetShootingDateTimeFromFileName(path);
                        if (!dateTime.HasValue)
                        {
                            // 撮影日時を取得できなければ無視
                            writer.WriteLine("Invalid DateTime  " + path);
                            continue;
                        }

                        // 同じ撮影日時のファイルをメディアプールから探す
                        string alternativeMediaPath = searchMedia.FindMedia(dateTime.Value);
                        if (alternativeMediaPath == "")
                        {
                            writer.WriteLine("Cannot find " + toString(dateTime, path));
                            continue;
                        }

                        // 見つかった
                        writer.WriteLine("Found " + path + " " + alternativeMediaPath);

                        pathPair.Add(path, alternativeMediaPath);
                    }
                }
            }

            foreach (Track track in vegas.Project.Tracks)
            {
                foreach (TrackEvent trackEvent in track.Events)
                {
                    if (trackEvent.Takes.Count != 1)
                    {
                        // 複数テイクあるイベントは無視
                        // （すでにMediaの置き換えが行われているので）
                        continue;
                    }

                    Take take = trackEvent.Takes[0];
                    {
                        Media media = take.Media;
                        string path = media.FilePath; // メディアのファイルパス

                        string alternativeMediaPath = "";
                        if (pathPair.TryGetValue(path, out alternativeMediaPath))
                        {
                            // 見つかったファイルをテイクに追加して、アクティブテイクにする
                            Media alternativeMedia = Media.CreateInstance(vegas.Project, alternativeMediaPath);

                            MediaStream mediaStream;

                            if (track.IsAudio())
                            {
                                mediaStream = alternativeMedia.GetAudioStreamByIndex(0);
                            }
                            else
                            {
                                mediaStream = alternativeMedia.GetVideoStreamByIndex(0);
                            }

                            Take alternativeTake = new Take(mediaStream);
                            trackEvent.Takes.Add(alternativeTake);
                            trackEvent.ActiveTake = alternativeTake;

                        }
                    }
                }
            }

                
            writer.Close();
            MessageBox.Show("終了しました。");
        }

        // ダイアログを開きファイルパスをユーザーに選択させる
        private string GetFilePath(string rootFilePath, string preFix)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FileName = preFix + System.IO.Path.GetFileNameWithoutExtension(rootFilePath) + ".txt";
            sfd.InitialDirectory = System.IO.Path.GetDirectoryName(rootFilePath) + "\\";
            sfd.Filter = "テキストファイル(*.txt)|*.txt";
            if (sfd.ShowDialog() != DialogResult.OK)
            {
                return "";
            }

            return sfd.FileName;
        }

        // 与えられたファイルパスの拡張子が、extentionsに含まれている拡張子であるかを調べる
        private bool isSameExtention(string path, string[] extentions)
        {
            // ファイルパスから拡張子を取得
            string searchExtention = Path.GetExtension(path).ToLower();

            // extentionsに含まれている拡張子かを調べる
            foreach (string extention in extentions)
            {
                if (extention == searchExtention)
                {
                    return true;
                }
            }

            return false;
        }

        // 画像のファイルパスかどうかを返す
        // 画像かどうかは、拡張子で判断する
        // HEIFの exif は読み方がわからないので、非対応
        private bool IsImage(string path)
        {
            // 画像の拡張子リスト
            string[] Extentions = new string[] { ".jpg", ".jpeg"};

            return isSameExtention(path, Extentions);
        }

        // 動画のファイルパスかどうかを返す
        // 動画かどうかは、拡張子で判断する
        private bool IsVideo(string path)
        {
            // 動画の拡張子リスト
            string[] Extentions = new string[] { ".mp4", ".mov" };

            return isSameExtention(path, Extentions);
        }

        private bool IsMOV(string path)
        {
            // MOV の拡張子リスト
            string[] Extentions = new string[] { ".mov" };

            return isSameExtention(path, Extentions);
        }

        private bool IsHEIF(string path)
        {
            // MOV の拡張子リスト
            string[] Extentions = new string[] { ".heic" };

            return isSameExtention(path, Extentions);
        }


        private string GetRecordedDateFromMOV(MediaInfo.MediaInfo mediaInfo)
        {
            string inform = mediaInfo.Inform();

            StringReader stringReader = new StringReader(inform);
            while (true)
            {
                string line = stringReader.ReadLine();
                if (line == null)
                {
                    break;
                }

                // 撮影日時は以下のように格納されている
                // com.apple.quicktime.creationdate         : 2021-07-31T18:34:03+0900

                if (!line.StartsWith("com.apple.quicktime.creationdate"))
                {
                    continue;
                }

                Match match = Regex.Match(line, "[0-9]");
                if (!match.Success)
                {
                    // ここまで見つかっておいて数値が見つからないのはおかしい
                    continue;
                }

                // 見つかった
                string recordedDate = line.Substring(match.Index);
                return recordedDate;
            }

            // みつからなかった
            return "";
        }

        // pathの動画のメタデータを読み込み、撮影日時を返す
        private DateTime? GetShootingDateTimeVideo(string path, System.IO.StreamWriter writer)
        {
            MediaInfo.MediaInfo mediaInfo = new MediaInfo.MediaInfo();

            IntPtr handle = mediaInfo.Open(path);
            if (handle == null)
            {

                writer.WriteLine("Cannot open");
            }

            // iPhoneで撮影した動画はデジカメで撮影した動画とでは見る場所を変える
            bool isUTC = false;
            string recordedDate = "";
            if (IsMOV(path))
            {
                recordedDate = GetRecordedDateFromMOV(mediaInfo);

                // "yyyy-MM-ddTHH:mm:ss+0900";
                // の形式になっている

                recordedDate = recordedDate.Replace('T', ' ');

                int indexPlus = recordedDate.IndexOf('+');
                if (0 <= indexPlus)
                {
                    // +があれば、UTCではない。 + 以降を削除する
                    isUTC = false;
                    recordedDate = recordedDate.Substring(0, indexPlus);
                }

                // "yyyy-MM-dd HH:mm:ss";
                // の形式になっている
            }
            else
            {
                // iPhoneで撮影した動画は Encoded_Date と完全一致しないことがある
                recordedDate = mediaInfo.Get(MediaInfo.StreamKind.General, 0, "Encoded_Date");

                // "UTC yyyy-MM-dd HH:mm:ss";
                // の形式になっている

                if (recordedDate.StartsWith("UTC "))
                {
                    isUTC = true;
                    recordedDate = recordedDate.Replace("UTC ", "");
                }

                // "yyyy-MM-dd HH:mm:ss";
                // の形式になっている
            }

            mediaInfo.Close();


            // 撮影日時を、コロン区切りにする
            recordedDate = recordedDate.Replace('-', ':').Replace(' ', ':');

            // 撮影日時を、コロン区切りで分割する
            string[] dividedDate = recordedDate.Split(':');

            // 分割した撮影日時から DateTime を取得する
            DateTime? dateTime = FromDividedDate(dividedDate);

            if (isUTC && dateTime.HasValue)
            {
                dateTime = dateTime?.ToLocalTime();
            }

            return dateTime;
        }

        private string GetShootingDateTimeImageUsingMediaInfo(string path, System.IO.StreamWriter writer)
        {
            MediaInfo.MediaInfo mediaInfo = new MediaInfo.MediaInfo();
            IntPtr handle = mediaInfo.Open(path);
            if (handle == null)
            {
                return "";
            }

            string inform = mediaInfo.Inform();
            writer.WriteLine("MediaInfo for " + path);
            writer.WriteLine(inform);

            return "";
        }

        // JPGは読めるが、HEIFは読めない
        private string GetShootingDateTimeImageUsingImage(string path)
        {
            System.Drawing.Image image = System.Drawing.Image.FromFile(path);
            System.Drawing.Imaging.PropertyItem propertyItem = image.GetPropertyItem(ExifTag.ShootingDateAndTime);

            string date = Encoding.ASCII.GetString(propertyItem.Value, 0, 19);

            image.Dispose();

            return date;
        }

        // JPGは読めるが、HEIFは読めない
        private string GetShootingDateTimeImageUsingBitmap(string path)
        {
            // Bitmap として読み込む
            System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(path);

            // 撮影日時を取得する
            int[] propertyIdList = bitmap.PropertyIdList;
            int index = Array.IndexOf(propertyIdList, ExifTag.ShootingDateAndTime);
            if (index == -1)
            {
                // 画像に撮影日時が含まれてない
                return null;
            }

            System.Drawing.Imaging.PropertyItem propertyItem = bitmap.PropertyItems[index];

            string date = Encoding.ASCII.GetString(propertyItem.Value, 0, 19);

            bitmap.Dispose();
            return date;
        }

        // pathの画像の Exif を読み込み、撮影日時を返す
        private DateTime? GetShootingDateTimeImage(string path)
        {
            // 撮影日時は、YYYY:MM:DD HH:MM:SS の形式で文字列として記録されている
            string date = GetShootingDateTimeImageUsingBitmap(path);
//            string date = GetShootingDateTimeImageUsingImage(path);
            if (date == null)
            {
                return null;
            }

            // 撮影日時を、コロン区切りにする
            date = date.Replace(' ', ':');

            // 撮影日時を、コロン区切りで分割する
            string[] dividedDate = date.Split(':');

            // 分割した撮影日時から DateTime を取得する
            DateTime? dateTime = FromDividedDate(dividedDate);
            return dateTime;
        }

        // みてねからダウンロードしたファイル名（YYYY-MM-DDTHHMMSS_comment.jpg）からYYYY-MM-DDTHHMMSSの部分を取り出す
        private string toDateTimeString(string path)
        {
            // フルパスからファイル名を取得する
            string filename = System.IO.Path.GetFileName(path);

            // ファイル名を _ で分割して、撮影日時とコメントに分割する
            string[] dateTimeAndComment = filename.Split('_');
            if (dateTimeAndComment.Length != 2)
            {
                return "";
            }

            return dateTimeAndComment[0];   // 撮影日時を返す
        }



        // pathの画像のファイル名から、撮影日時を返す
        // みてねからダウンロードしたファイルは、YYYY-MM-DDTHHMMSS_comment.jpg のようになっている
        private DateTime? GetShootingDateTimeFromFileName(string path)
        {
            // ファイルパスから撮影日時を表す文字列を取得する
            string dateTimeString = toDateTimeString(path);

            // dateTimeString として YYYY-MM-DDTHHMMSS が得られる

            // YYYY-MM-DDTHHMMSS を T で分割して Date と Time を表す文字列を得る
            string[] dateAndTime = dateTimeString.Split('T');
            if (dateAndTime.Length != 2)
            {
                return null;
            }

            // ここまでで以下のように分割できている：
            // dateAndTime[0]：YYYY-MM-DD
            // dateAndTime[1]: HHMMSS

            string date = dateAndTime[0];

            // date を - で分割する
            string[] dateElements = date.Split('-');
            if (dateElements.Length != 3)
            {
                return null;
            }

            string time = dateAndTime[1];
            if (time.Length != 6)
            {
                return null;
            }

            // 分割した各要素を文字列配列にコピーする
            string[] dividedDate = new string[6];
            dividedDate[0] = dateElements[0];       // year
            dividedDate[1] = dateElements[1];       // month
            dividedDate[2] = dateElements[2];       // day
            dividedDate[3] = time.Substring(0, 2);  // hour
            dividedDate[4] = time.Substring(2, 2);  // minute
            dividedDate[5] = time.Substring(4, 2);  // second

            // 分割した撮影日時から DateTime を取得する
            DateTime? dateTime = FromDividedDate(dividedDate);
            return dateTime;
        }

        // 日時を YYYY:MM:DD:HH:MM:SS の 6 つに分割した文字列配列から、DateTimeを作成する
        private DateTime? FromDividedDate(string[] dividedDate)
        {
            if (dividedDate.Length != 6)
            {
                // 形式が違えばエラーとする
                return null;
            }

            try
            {
                int year = Int32.Parse(dividedDate[0]);
                int month = Int32.Parse(dividedDate[1]);
                int day = Int32.Parse(dividedDate[2]);
                int hour = Int32.Parse(dividedDate[3]);
                int minute = Int32.Parse(dividedDate[4]);
                int second = Int32.Parse(dividedDate[5]);

                DateTime dateTime = new DateTime(year, month, day, hour, minute, second);
                return dateTime;
            }
            catch (FormatException)
            {
            }

            return null;
        }


        private string toString(DateTime? dateTime, string path)
        {
            if (dateTime == null)
            {
                return path;
            }

            return dateTime?.ToString("yyyy/MM/dd HH:mm:ss") + " " + path;
        }

    }

}
