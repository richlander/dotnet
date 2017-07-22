using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Console;

namespace ConsoleApp4
{
    class Program
    {
        static void Main(string[] args)
        {
            var path = @"C:\Users\rlander\Documents\5d.csv";
            var includeSecurity = false;

            var releases = new List<Release>();

            using (var reader = new StreamReader(path))
            {
                var line = string.Empty;
                while ((line = reader.ReadLine()) != null)
                {
                    var release = new Release();
                    releases.Add(release);

                    var segments = line.Split(',');
                    // OS,Parent KB,Child KBs,Parent KB,Child KBs

                    var product = GetWindowsProduct(segments[0]);
                    release.Product = product;

                    var qkb = new KB();
                    qkb.Product = "QR";
                    qkb.Number = segments[1];
                    qkb.Children = GetNETKBs(segments[2]);
                    release.Childen.Add(qkb);

                    if (segments.Length> 3)
                    {
                        includeSecurity = true;
                        var skb = new KB();
                        skb.Product = "SR";
                        skb.Number = segments[3];
                        skb.Children= GetNETKBs(segments[4]);
                        release.Childen.Add(skb);
                    }
                }
            }

            PrintKBs(releases, includeSecurity);


        }

        static void PrintKBs(List<Release> releases, bool securityRelease)
        {
            const string kblink = "https://support.microsoft.com/kb/";
            const string cataloglink = "http://www.catalog.update.microsoft.com/Search.aspx?q=";
            const string notApplicable = "N/A";

            releases = releases.OrderByDescending(r => r.Product.Weight).ToList();

            WriteLine("<table>");
            WriteLine("<thead><tr>");
            WriteLine("<th>Product Version</th><th>Security and Quality Rollup KB</th>");
            if (securityRelease)
            {
                WriteLine("<th>Security Rollup KB</th>");
            }
            WriteLine("</tr></thead>");

            foreach (var release in releases)
            {
                // http://www.catalog.update.microsoft.com/Search.aspx?q=4014983
                // https://support.microsoft.com/kb/4014983

                var qkb = release.Childen[0];
                var skb = release.Childen.Count > 1 ? release.Childen[1] : null;

                var qualityRow = string.IsNullOrWhiteSpace(qkb.Number) ? notApplicable : $"<td><strong><a href =\"{cataloglink}{release.Childen[0].Number}\">Catalog</a><BR><a href=\"https://support.microsoft.com/kb/{qkb.Number}\">{qkb.Number}</a></strong></td>";
                var securityRow = string.Empty;

                if (!securityRelease)
                {
                }
                else if (string.IsNullOrWhiteSpace(skb?.Number))
                {
                    securityRow = notApplicable;
                }
                else
                {
                    securityRow = $"<td><strong><a href =\"{cataloglink}{release.Childen[1].Number}\">Catalog</a><BR><a href=\"https://support.microsoft.com/kb/{skb.Number}\">{skb.Number}</a></strong></td>";
                }
                

                WriteLine("<tr>");
                WriteLine($"<td><strong>{release.Product.Product}</strong></td>"+
                          $"{qualityRow}" +
                          $"{securityRow}");
                WriteLine("</tr>");
                WriteLine("<tr>");

                for (int i =0; i < qkb.Children.Count;i++)
                {
                    var qualityRowForProduct = (qkb.Children.Count <= i || string.IsNullOrEmpty(qkb.Children[i].Number)) ? string.Empty : $"<td><a href=\"{kblink}{qkb.Children[i].Number}\">{qkb.Children[i].Number}</a></td>";
                    var securityRowForProduct = (skb == null || skb.Children.Count <= i || string.IsNullOrEmpty(skb.Children[i].Number)) ? string.Empty : $"<td><a href=\"{kblink}{skb.Children[i].Number}\">{skb.Children[i].Number}</a></td>";

                    WriteLine("<tr>");
                    WriteLine($"<td style=\"padding-left:.5cm\">{qkb.Children[i].Product}</td>"+
                              $"{qualityRowForProduct}"+
                              $"{securityRowForProduct}");
                    WriteLine("</tr>");
                }
                
            }
            WriteLine("</table>");
        }

        static List<KB> GetNETKBs(string kbstring)
        {
            var kbs = new List<KB>();
            var segments = kbstring.Split(' ');

            for (var i =0; i < segments.Length;i++)
            {
                var segment = segments[i];
                if (segment.Contains('.'))
                {
                    var kb = new KB();
                    kbs.Add(kb);
                    var buffer = new StringBuilder();
                    buffer.Append(".NET Framework ");
                    var versions = segment.Split('/');

                    for (var j = 0; j<versions.Length;j++)
                    {
                        var version = versions[j];
                        buffer.Append(version);
                        if (j+1 < versions.Length)
                        {
                            buffer.Append(", ");
                        }
                    }

                    kb.Product = buffer.ToString();
                    i++;
                    i++;
                    kb.Number = segments[i];
                }
            }

            kbs.Reverse();

            return kbs;
        }

        static WeightedProduct GetWindowsProduct(string name)
        {
            var products = new List<WeightedProduct>();
            var segments = name.Split('/');

            var canonicalNames = new Dictionary<string, WeightedProduct>
            {
                { "Win 10 update 1703", new WeightedProduct(15,"Windows 10 Creators Update") },
                { "Win 10 update 1607", new WeightedProduct(14,"Windows 10 Anniversary Update") },
                { "Server 2016", new WeightedProduct(13,"Windows Server 2016") },
                { "Win 10 update 1511", new WeightedProduct(12,"Windows 10 1511") },
                { "Win 10 RTM", new WeightedProduct(11,"Windows 10 1507") },
                { "Win 8.1", new WeightedProduct(10,"Windows 8.1") },
                { "Server 2k12 R2", new WeightedProduct(9,"Windows Server 2012 R2")},
                { "Server 2k12", new WeightedProduct(8,"Windows Server 2012")},
                { "Win 7", new WeightedProduct(7,"Windows 7")},
                { "Vista", new WeightedProduct(6,"Windows Vista")},
                { "Server 2k8 R2", new WeightedProduct(5,"Windows Server 2008 R2")},
                { "Server 2k8", new WeightedProduct(4,"Windows Server 2008")}
            };

            foreach (var s in segments)
            {
                var prod = s;
                prod = prod.Trim();
                var product = canonicalNames[prod];
                products.Add(product);
            }

            products = products.OrderByDescending(x => x.Weight).ToList();

            var buffer = new StringBuilder();
            for (var i = 0; i < products.Count;i++)
            {
                buffer.Append(products[i].Product);

                if (i+1 <products.Count)
                {
                    buffer.Append("<BR>");
                }
            }

            var weightedProduct = new WeightedProduct(products[0].Weight, buffer.ToString());

            return weightedProduct;
        }
    }

    public class WeightedProduct
    {
        public WeightedProduct(int weight, string product)
        {
            Weight = weight;
            Product = product;
        }

        public int Weight;
        public string Product;
    }

    public class Release
    {
        public WeightedProduct Product;
        public List<KB> Childen = new List<KB>();
    }


    public class KB
    {
        public string Number;
        public string Product;
        public List<KB> Children = new List<KB>();
    }
}
