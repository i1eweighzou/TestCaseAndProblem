using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TestCaseAndProblem
{
    public class PictureUtils
    {
        public static String get_picture_path(String guid_str = "")
        {
            if (String.IsNullOrEmpty(guid_str)) {
                guid_str = Guid.NewGuid().ToString("N");
            }

            String dir = Globals.EditItemsSheet.get_file_save_dir() + @"picture\";
            if (!Directory.Exists(dir)) {
                Directory.CreateDirectory(dir);
            }
            return dir + guid_str + ".JPG";
        }

        public static String insert_picture()
        {
            String guid_str = "";
            IDataObject iData = Clipboard.GetDataObject();
            if (iData.GetDataPresent(DataFormats.MetafilePict))
            {
                Image img = Clipboard.GetImage();
                if (img != null)
                {
                    guid_str = save_image(img);
                }
            }
            else if (iData.GetDataPresent(DataFormats.Bitmap))
            {
                Image img = iData.GetData(DataFormats.Bitmap) as Image;
                if (img != null)
                {
                    guid_str = save_image(img);
                }
            }
            else if (iData.GetDataPresent(DataFormats.FileDrop))
            {
                var files = Clipboard.GetFileDropList();
                if (files.Count == 0) { return ""; }
                Image img = Image.FromFile(files[0]);
                if (img != null)
                {
                    guid_str = save_image(img);
                }
            }
            else if (iData.GetDataPresent(DataFormats.Text))
            {
                var path = (String)iData.GetData(DataFormats.Text);
                var chars = Path.GetInvalidPathChars();
                if (path.IndexOfAny(chars) >= 0)
                {
                    return "";
                }
                if (System.IO.File.Exists(path))
                {
                    var name = Path.GetFileNameWithoutExtension(path);
                    var extension = path.Substring(path.LastIndexOf("."));
                    string imgType = ".png|.jpg|.jpeg";
                    if (imgType.Contains(extension.ToLower()))
                    {
                        Image img = Image.FromFile(path); ;
                        if (img != null)
                        {
                            guid_str = save_image(img);
                        }
                    }
                }
            }

            return guid_str;
        }

        private static string save_image(Image img)
        {
            string guid_str = Guid.NewGuid().ToString("N");
            img.Save(get_picture_path(guid_str));
            return guid_str;
        }

        public static Image getImage(String guid_str) {
            return Image.FromFile(get_picture_path(guid_str));
        }
    }
}
