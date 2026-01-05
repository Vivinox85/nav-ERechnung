using s2industries.ZUGFeRD;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ERechnung.Models
{
    internal class XRechnung
    {
        public string InvoiceNumber { get; set; }
        public DateTime InvoiceDate { get; set; }
        public CurrencyCodes Currency { get; set; }
        public string OrderNumber { get; set; }
        public DateTime DeliveryDate { get; set; }
        public string PaymentTerms { get; set; }
        public DateTime PaymentDueDate { get; set; }
        public Seller Seller { get; set; }
        public Buyer Buyer { get; set; }
        public Party DeliveryAddress { get; set; }
        public List<LineItem> LineItems { get; set; }
        public List<Bankkonto> BankAccounts { get; set; }
        public List<PaymentTerms> SkontoOptions {  get; set; }
        public decimal TotalNetAmount
        {
            get
            {
                return LineItems.Sum(x => x.LineTotal);
            }
            set
            {
                TotalNetAmount = value;
            }
        }
        public decimal TotalTaxAmount
        {
            get
            {
                return LineItems.Sum(x => x.TaxAmount);
            }
            set
            {
                TotalTaxAmount = value;
            }
        }
        public decimal TotalGrossAmount
        {
            get
            {
                return LineItems.Sum(x => x.Total);
            }
            set
            {
                TotalGrossAmount = value;
            }
        }
        public decimal TotalAllowanceChargeAmount { get; set; }
        public decimal TotalChargeAmount { get; set; }
        public decimal TotalPrepaidAmount { get; set; }
        public decimal TotalRoundingAmount { get; set; }
        public decimal TotalDueAmount
        {
            get
            {
                return TotalGrossAmount - TotalPrepaidAmount + TotalTaxAmount;
            }
        }

        // create XML from XRechnung object with s2industries.ZUGFeRD
        public bool CreateXML(string filePath)
        {
            bool success = false;
            TaxCategoryCodes overallTaxCategory = TaxCategoryCodes.S;
            decimal overallTaxPercent = 19m;

            InvoiceDescriptor desc = InvoiceDescriptor.CreateInvoice(invoiceNo:this.InvoiceNumber, invoiceDate:this.InvoiceDate, currency:this.Currency);
            desc.ReferenceOrderNo = this.OrderNumber;
            
            // Verwendungszweck für Zahlung:
            desc.PaymentReference = this.InvoiceNumber;

            desc.SetBuyer(name:this.Buyer.Name, postcode:this.Buyer.ZipCode, city:this.Buyer.City, street:this.Buyer.Street, country:this.Buyer.Country, id:this.Buyer.ID);
            desc.AddBuyerTaxRegistration(no:this.Buyer.VATID, schemeID:TaxRegistrationSchemeID.VA);
            desc.SetBuyerContact(name:this.Buyer.Contact, emailAddress:this.Buyer.Email);
            desc.SetBuyerOrderReferenceDocument(orderNo:this.Buyer.OrderReferenceDocument, orderDate:this.Buyer.OrderReferenceDocumentDate);
            desc.SetBuyerElectronicAddress(address:this.Buyer.Email, electronicAddressSchemeID:ElectronicAddressSchemeIdentifiers.EM);

            desc.SetSeller(name:this.Seller.Name, postcode:this.Seller.ZipCode, city:this.Seller.City, street:this.Seller.Street, country:this.Seller.Country, id: this.Seller.ID, description:this.Seller.LegalDescription);
            desc.AddSellerTaxRegistration(no:this.Seller.VATID, schemeID:TaxRegistrationSchemeID.VA);
            desc.AddSellerTaxRegistration(no: this.Seller.TaxNumber, schemeID: TaxRegistrationSchemeID.FC);
            desc.SetSellerContact(name:this.Seller.Contact, orgunit:this.Seller.OrganizationUnit, emailAddress:this.Seller.Email, phoneno:this.Seller.Phone);
            desc.SetSellerElectronicAddress(address:this.Seller.Email, electronicAddressSchemeID:ElectronicAddressSchemeIdentifiers.EM);

            desc.ShipTo = DeliveryAddress;

            desc.ActualDeliveryDate = this.DeliveryDate;

            desc.SetTotals(
                lineTotalAmount:this.TotalNetAmount,
                chargeTotalAmount:this.TotalChargeAmount,
                allowanceTotalAmount:this.TotalAllowanceChargeAmount,
                taxBasisAmount:this.TotalGrossAmount,
                taxTotalAmount:this.TotalTaxAmount,
                grandTotalAmount:this.TotalNetAmount + this.TotalTaxAmount,
                totalPrepaidAmount:this.TotalPrepaidAmount,
                duePayableAmount:this.TotalDueAmount);

            desc.AddTradePaymentTerms(description:this.PaymentTerms, dueDate:this.PaymentDueDate);

            foreach(PaymentTerms pt in this.SkontoOptions)
            {
                desc.AddTradePaymentTerms(pt.Description, pt.DueDate, pt.PaymentTermsType, pt.DueDays, pt.Percentage);
            }
            
            foreach (LineItem lineItem in this.LineItems)
            {
                TradeLineItem curItem = desc.AddTradeLineItem(name:lineItem.Name, netUnitPrice:lineItem.UnitPrice, unitCode:lineItem.Unit, description:lineItem.Description, billedQuantity:lineItem.Quantity, grossUnitPrice:lineItem.UnitPrice + (lineItem.UnitPrice * lineItem.TaxPercent / 100), lineTotalAmount:lineItem.LineTotal, taxType:lineItem.TaxType, categoryCode:lineItem.TaxCategory, taxPercent:lineItem.TaxPercent, sellerAssignedID:lineItem.ID, buyerAssignedID:lineItem.CustomerID);
                curItem.OriginTradeCountry = lineItem.OriginCountry;
                if (lineItem.TaxCategory == TaxCategoryCodes.Z)
                {
                    overallTaxCategory = TaxCategoryCodes.Z;
                }
                if(lineItem.TaxPercent == 0)
                {
                    overallTaxPercent = 0m;
                }
            }

            desc.AddApplicableTradeTax(basisAmount:this.TotalNetAmount, percent:overallTaxPercent, taxAmount:this.TotalTaxAmount, typeCode:TaxTypes.VAT, categoryCode:overallTaxCategory);

            desc.SetPaymentMeans(paymentCode:PaymentMeansTypeCodes.SEPACreditTransfer);

            foreach (Bankkonto bankkonto in this.BankAccounts)
            {
                desc.AddCreditorFinancialAccount(iban: bankkonto.IBAN, bic: bankkonto.BIC, bankleitzahl: bankkonto.Bankleitzahl, bankName: bankkonto.Bankname, name: bankkonto.Kontoinhaber);
            }            

            // X-Rechnung:
            desc.BusinessProcess = "urn:fdc:peppol.eu:2017:poacc:billing:01:1.0";
          
            FileStream stream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
            desc.Save(stream:stream, version:ZUGFeRDVersion.Version23, profile:Profile.XRechnung);
            stream.Flush();
            stream.Close();
            success = true;
            return success;
        }
    }

    public class LineItem
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string CustomerID { get; set; }
        public QuantityCodes Unit { get; set; }
        public decimal Quantity { get; set; }
        public decimal UnitPrice { get; set; }
        public decimal LineTotal { get; set; }
        /*{
            get
            {
                return Quantity * UnitPrice;
            }
            set
            {
                LineTotal = value;
            }
        }*/
        public decimal TaxAmount
        {
            get
            {
                return LineTotal * (TaxPercent / 100);
            }
            set
            {
                TaxAmount = value;
            }
        }
        public decimal Total
        {
            get
            {
                return LineTotal;
            }
            set
            {
                Total = value;
            }
        }
        public TaxTypes TaxType { get; set; }
        public TaxCategoryCodes TaxCategory { get; set; }
        public decimal TaxPercent { get; set; }
        public CountryCodes OriginCountry { get; set; }
    }

    public class Buyer
    {
        public string Name { get; set; }
        public string ZipCode { get; set; }
        public string City { get; set; }
        public string Street { get; set; }
        public CountryCodes Country { get; set; }
        public string VATID { get; set; }
        public TaxRegistrationSchemeID TaxRegistrationSchemeID { get; set; }
        public string Contact { get; set; }
        public string OrganizationUnit { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string ID { get; set; }
        public string OrderReferenceDocument { get; set; }
        public DateTime OrderReferenceDocumentDate { get; set; }
    }

    public class Seller : Buyer
    {
        public string TaxNumber { get; set; }
        public string TaxNumberType { get; set; }
        public string LegalDescription { get; set; }
    }

    public class Bankkonto
    {
        public string IBAN {  get; set; }
        public string BIC {  get; set; }
        public string Bankleitzahl { get; set; }
        public string Bankname { get; set; }
        public string Kontoinhaber { get; set; }        
    }
}
