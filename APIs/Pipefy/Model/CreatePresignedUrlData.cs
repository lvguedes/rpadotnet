namespace RpaLib.APIs.Pipefy.Model
{
    public class CreatePresignedUrlData
    {
        public string ClientMutationId { get; set; }
        public string DownloadUrl { get; set; }
        public string Url { get; set; }
    }
}