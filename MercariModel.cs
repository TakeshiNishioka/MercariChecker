using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MercariChecker
{
    public class MercariModel
    {
        public MercariModel()
        {
            Items = new List<ItemModel>();
        }

        public string Keywords { set; get; }

        public List<ItemModel> Items;
    };

    public class ItemModel
    {
        public string Name { set; get; }
        public string DetailUrl { set; get; }
        public string ImgUrl { set; get; }
        public string Price { set; get; }
        public string Keyword { set; get; }
    };

}
