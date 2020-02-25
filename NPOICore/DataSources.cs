using System.Collections.Generic;
using System.Linq;

namespace NPOICore
{
    public class DataSource
    {
        public string Code { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
    }

    public static class DataSources
    {
        public static List<DataSource> AreaManagers
        {
            get
            {
                return Enumerable.Range(0, 10).Select(x => new DataSource
                {
                    Code = "AMCode " + x,
                    LastName = "AMLast Name " + x,
                    FirstName = "AMName " + x
                }).ToList();
            }
        }

        public static List<DataSource> Sales
        {
            get
            {
                return Enumerable.Range(0, 10).Select(x => new DataSource
                {
                    Code = "SLCode " + x,
                    LastName = "SLLast Name " + x,
                    FirstName = "SLName " + x
                }).ToList();
            }
        }

        public static List<DataSource> Products
        {
            get
            {
                return Enumerable.Range(0, 10).Select(x => new DataSource
                {
                    Code = "PRCode " + x,
                    LastName = "PRLast Name " + x,
                    FirstName = "PRName " + x
                }).ToList();
            }
        }
    }
}
