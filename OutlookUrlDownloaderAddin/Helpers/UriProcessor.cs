using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using Microsoft.Office.Interop.Outlook;

namespace OutlookUrlDownloaderAddin.Helpers
{
    public interface IUriProcessor
    {
        Task<ParsedObject> ParseAsync(MailItem mailItem);
    }

    public class UriProcessor : IUriProcessor
    {
        public UriProcessor()
        {

        }

        public Task<ParsedObject> ParseAsync(MailItem mailItem)
        {
            var task = Task.Factory.StartNew<ParsedObject>(() =>
            {
                return Parse(mailItem);
            });

            return task;
        }

        private ParsedObject Parse(MailItem mailItem)
        {
            var parsedObject = new ParsedObject();

            if (mailItem == null)
                return parsedObject;

            //https://onedrive.live.com/redir?page=view&resid=68D3D159D743480C!121606&authkey=!AOAOwU50zyPBPDM                     
            var linkParser = new Regex(@"https://[^\s""<\(\)]+", RegexOptions.Compiled | RegexOptions.IgnoreCase);

            foreach (var match in linkParser.Matches(mailItem.HTMLBody))
            {
                var urlString = match.ToString();

                if (!urlString.Contains(UrlValueNames.BaseUrl))
                    continue;

                urlString = urlString.Replace("&amp;", "&");

                var parsedUrlObject = new ParsedUrlObject(urlString, HttpUtility.ParseQueryString(urlString.Substring(urlString.IndexOf("?"))));
                parsedObject.UrlObjects.Add(parsedUrlObject);

                Debug.WriteLine("{0} : {1}", UrlValueNames.ResId, parsedUrlObject.ParsedQueryData.Get(UrlValueNames.ResId));
                Debug.WriteLine("{0} : {1}", UrlValueNames.AuthKey, parsedUrlObject.ParsedQueryData.Get(UrlValueNames.AuthKey));
            }

            return parsedObject;
        }
    }

    public class ParsedObject
    {
        private readonly List<ParsedUrlObject> urlObjects;

        public ParsedObject()
        {
            urlObjects = new List<ParsedUrlObject>();
        }

        public List<ParsedUrlObject> UrlObjects
        {
            get
            {
                return urlObjects;
            }
        }

        public bool HasUrls
        {
            get
            {
                return urlObjects.Count > 0;
            }
        }
    }

    public class ParsedUrlObject
    {
        private string url;
        private NameValueCollection parsedQueryData;

        public ParsedUrlObject(string url, NameValueCollection parsedQueryData)
        {
            this.url = url;
            this.parsedQueryData = parsedQueryData;
        }

        public string Url
        {
            get
            {
                return url;
            }
        }

        public NameValueCollection ParsedQueryData
        {
            get
            {
                return parsedQueryData;
            }
        }
    }

    public static class UrlValueNames
    {
        public const string BaseUrl = "onedrive.live.com";
        public const string ResId = "resid";
        public const string AuthKey = "authkey";
    }
}