﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace locationapi.Models
{
    public class ImportServiceOrderModel
    {
        public string StoreNo { get; set; }
        public DateTime? OrderCreationDate { get; set; }
        public string DocumentNo { get; set; }
        public string ServiceOrderNo { get; set; }
        public string ServiceItemNo { get; set; }
        public string ServiceName { get; set; }
        public DateTime? ServiceDate { get; set; }
        public string ServiceTimeSlot { get; set; }
        public string ServiceStatus { get; set; }
        public decimal? ServiceGoodsValue { get; set; }
        public string CapacityUnit { get; set; }
        public decimal? CapacityValueWeight { get; set; }
        public decimal? CapacityValueVolume { get; set; }
        public decimal? BookedQty { get; set; }
        public decimal? ServicePriceExclVAT { get; set; }
        public decimal? ServicePriceInclVAT { get; set; }
        public string PriceCalcMethod { get; set; }
        public decimal? NoofItems { get; set; }
        public decimal? NoofPackages { get; set; }
        public decimal? TotalOrderValue { get; set; }
        public string ServiceProviderName { get; set; }
        public string ServiceProviderID { get; set; }
        public string PaymentStatus { get; set; }
        public string PaymenttoIKEA_SP { get; set; }
        public string ShipToCustomerName { get; set; }
        public string ShipToAddress { get; set; }
        public string ShipToAddress2 { get; set; }
        public string ShipToPostcode { get; set; }
        public string ShipToCity { get; set; }
        public string ShipToPhoneNo { get; set; }
        public string ShipToEmail { get; set; }
        public string SellToCustomerName { get; set; }
        public string SellToAddress { get; set; }
        public string SellToAddress2 { get; set; }
        public string SellToPostcode { get; set; }
        public string SellToCity { get; set; }
        public string SellToPhoneNo { get; set; }
        public string SellToMobilePhoneNo { get; set; }
        public string SellToEmail { get; set; }
        public string ServiceComment { get; set; }
        public string OrderComment { get; set; }
        public string SalesPerson { get; set; }
        public string CRMCaseID { get; set; }
        public DateTime? HandoverDate { get; set; }
        public DateTime? HandoverTime { get; set; }
    }
}
