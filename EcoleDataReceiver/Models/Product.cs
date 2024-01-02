using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EcoleDataReceiver.Models
{
    [Table("M商品")]
    public class Product
    {
        [Column("商品コード")]
        public string Id { get; set; }

        [Column("商品名")]
        public string ProductName { get; set; }

        [Column("削除区分")]
        public int State { get; set; }

        [Column("諸口区分")]
        public int Sundry { get; set; }

        [Column("課税区分")]
        public int TaxationType { get; set; }

        [Column("商品分類コード")]
        public int ProductCategoryId { get; set; }

        [Column("単位")]
        public string Unit { get; set; }

        [Column("売価")]
        public decimal Price { get; set; }

        [Column("原価")]
        public decimal Cost { get; set; }

        [Column("定価")]
        public decimal CatalogPrice { get; set; }

        [Column("在庫")]
        public decimal StockType { get; set; }

        [Column("JANコード")]
        public int JAN { get; set; }

        [Column("予備Number1")]
        public int ReserveNo1 { get; set; }

        [Column("予備Number2")]
        public int ReserveNo2 { get; set; }

        [Column("品番")]
        public string ProdutNo { get; set; }

        [Column("メーカー名")]
        public string MakerName { get; set; }

        [Column("読取専用区分")]
        public int IsReadOnly { get; set; }

        [Column("最終更新日時")]
        public decimal UpdatedAt { get; set; }
    }
}
