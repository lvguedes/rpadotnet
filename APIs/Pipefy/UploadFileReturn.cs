using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy
{
    public class UploadFileReturn
    {
        public string UploadUrl { get; private set; }
        public string DownloadUrl { get; private set; }
        public UploadFileReturn(string uploadUrl, string downloadUrl)
        {
            UploadUrl = uploadUrl;
            DownloadUrl = downloadUrl;
        }
    }
}
