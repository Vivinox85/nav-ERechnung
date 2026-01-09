using s2industries.ZUGFeRD;
using s2industries.ZUGFeRD.PDF;
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
        private InvoiceDescriptor desc { get; set; }
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
        public List<PaymentTerms> SkontoOptions { get; set; }
        public List<Note> Notes { get; set; }
        public decimal TotalNetAmount
        {
            get
            {
                return LineItems.Sum(x => x.LineTotal);
            }
        }
        public decimal TotalTaxAmount
        {
            get
            {
                return LineItems.Sum(x => x.TaxAmount);
            }
        }
        public decimal TotalGrossAmount
        {
            get
            {
                return LineItems.Sum(x => x.Total);
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
            bool success = true;
            try
            {
                FillInvoiceDescriptor();
                FileStream stream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
                desc.Save(stream: stream, version: ZUGFeRDVersion.Version23, profile: Profile.XRechnung);
                stream.Flush();
                stream.Close();
            }
            catch (Exception ex)
            {
                LogError(ex);                
                success = false;
            }
            return success;
        }

        public bool CreatePDF(string inPDFPath, string outPDFPath)
        {
            bool success = true;
            try
            {
                FillInvoiceDescriptor();
                InvoicePdfProcessor.SaveToPdf(outPDFPath, ZUGFeRDVersion.Version23, Profile.XRechnung, ZUGFeRDFormats.CII, inPDFPath, this.desc);
            }
            catch (Exception ex)
            {
                LogError(ex);
                success = false;
            }
            return success;
        }

        private void LogError(Exception ex)
        {
            try
            {
                string logPath = @"C:\Temp\ERechnung_Error.txt";

                if (!Directory.Exists(@"C:\Temp")) Directory.CreateDirectory(@"C:\Temp");

                using (StreamWriter sw = new StreamWriter(logPath, true))
                {
                    sw.WriteLine("--- " + DateTime.Now.ToString("g") + " ---");
                    sw.WriteLine("Message: " + ex.Message);
                    sw.WriteLine("StackTrace: " + ex.StackTrace);
                    if (ex.InnerException != null)
                    {
                        sw.WriteLine("Inner Exception: " + ex.InnerException.Message);
                        sw.WriteLine("Inner StackTrace: " + ex.InnerException.StackTrace);
                    }
                    sw.WriteLine("------------------------------------------");
                    sw.WriteLine();
                }
            }
            catch
            {                
            }
        }

        private void FillInvoiceDescriptor()
        {
            TaxCategoryCodes overallTaxCategory = TaxCategoryCodes.S;
            decimal overallTaxPercent = 19m;

            desc = InvoiceDescriptor.CreateInvoice(invoiceNo: this.InvoiceNumber, invoiceDate: this.InvoiceDate, currency: this.Currency);
            desc.ReferenceOrderNo = this.OrderNumber;

            // Verwendungszweck für Zahlung:
            desc.PaymentReference = this.InvoiceNumber;

            desc.SetBuyer(name: this.Buyer.Name, postcode: this.Buyer.ZipCode, city: this.Buyer.City, street: this.Buyer.Street2, receiver: this.Buyer.Street, country: this.Buyer.Country, id: this.Buyer.ID);
            desc.AddBuyerTaxRegistration(no: this.Buyer.VATID, schemeID: TaxRegistrationSchemeID.VA);
            desc.SetBuyerContact(name: this.Buyer.Contact, emailAddress: this.Buyer.Email);
            desc.SetBuyerOrderReferenceDocument(orderNo: this.Buyer.OrderReferenceDocument, orderDate: this.Buyer.OrderReferenceDocumentDate);
            desc.SetBuyerElectronicAddress(address: this.Buyer.Email, electronicAddressSchemeID: ElectronicAddressSchemeIdentifiers.EM);

            desc.SetSeller(name: this.Seller.Name, postcode: this.Seller.ZipCode, city: this.Seller.City, street: this.Seller.Street, country: this.Seller.Country, id: this.Seller.ID);
            desc.AddSellerTaxRegistration(no: this.Seller.VATID, schemeID: TaxRegistrationSchemeID.VA);
            desc.AddSellerTaxRegistration(no: this.Seller.TaxNumber, schemeID: TaxRegistrationSchemeID.FC);
            desc.SetSellerContact(name: this.Seller.Contact, orgunit: this.Seller.OrganizationUnit, emailAddress: this.Seller.Email, phoneno: this.Seller.Phone);
            desc.SetSellerElectronicAddress(address: this.Seller.Email, electronicAddressSchemeID: ElectronicAddressSchemeIdentifiers.EM);

            desc.ShipTo = DeliveryAddress;

            desc.ActualDeliveryDate = this.DeliveryDate;

            desc.SetTotals(
                lineTotalAmount: this.TotalNetAmount,
                chargeTotalAmount: this.TotalChargeAmount,
                allowanceTotalAmount: this.TotalAllowanceChargeAmount,
                taxBasisAmount: this.TotalGrossAmount,
                taxTotalAmount: this.TotalTaxAmount,
                grandTotalAmount: this.TotalNetAmount + this.TotalTaxAmount,
                totalPrepaidAmount: this.TotalPrepaidAmount,
                duePayableAmount: this.TotalDueAmount);

            desc.AddTradePaymentTerms(description: this.PaymentTerms, dueDate: this.PaymentDueDate);

            foreach (PaymentTerms pt in this.SkontoOptions)
            {
                desc.AddTradePaymentTerms(pt.Description, pt.DueDate, pt.PaymentTermsType, pt.DueDays, pt.Percentage);
            }

            foreach (Note curNote in this.Notes)
            {
                desc.AddNote(note: curNote.Content, subjectCode: curNote.SubjectCode);
            }

            foreach (LineItem lineItem in this.LineItems)
            {
                TradeLineItem curItem = desc.AddTradeLineItem(lineID: lineItem.ID, name: lineItem.Name, netUnitPrice: lineItem.UnitPrice, unitCode: lineItem.Unit, unitQuantity: lineItem.UnitQuantity, description: lineItem.Description, billedQuantity: lineItem.Quantity, grossUnitPrice: lineItem.UnitPrice + (lineItem.UnitPrice * lineItem.TaxPercent / 100), lineTotalAmount: lineItem.LineTotal, taxType: lineItem.TaxType, categoryCode: lineItem.TaxCategory, taxPercent: lineItem.TaxPercent, sellerAssignedID: lineItem.ID, buyerAssignedID: lineItem.CustomerID);
                curItem.OriginTradeCountry = lineItem.OriginCountry;
                if (lineItem.TaxCategory == TaxCategoryCodes.Z)
                {
                    overallTaxCategory = TaxCategoryCodes.Z;
                }
                if (lineItem.TaxPercent == 0)
                {
                    overallTaxPercent = 0m;
                }
            }

            desc.AddApplicableTradeTax(basisAmount: this.TotalNetAmount, percent: overallTaxPercent, taxAmount: this.TotalTaxAmount, typeCode: TaxTypes.VAT, categoryCode: overallTaxCategory);

            desc.SetPaymentMeans(paymentCode: PaymentMeansTypeCodes.SEPACreditTransfer);

            foreach (Bankkonto bankkonto in this.BankAccounts)
            {
                desc.AddCreditorFinancialAccount(iban: bankkonto.IBAN, bic: bankkonto.BIC, bankleitzahl: bankkonto.Bankleitzahl, bankName: bankkonto.Bankname, name: bankkonto.Kontoinhaber);
            }

            // X-Rechnung:
            desc.BusinessProcess = "urn:fdc:peppol.eu:2017:poacc:billing:01:1.0";

        }
    }

    public class LineItem
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string CustomerID { get; set; }
        public QuantityCodes Unit { get; set; }
        public decimal UnitQuantity { get; set; }
        public decimal Quantity { get; set; }
        public decimal UnitPrice { get; set; }
        public decimal LineTotal { get; set; }

        public decimal TaxAmount
        {
            get
            {
                return LineTotal * (TaxPercent / 100);
            }
        }
        public decimal Total
        {
            get
            {
                return LineTotal;
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
        public string Street2 { get; set; }
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
    }

    public class Bankkonto
    {
        public string IBAN { get; set; }
        public string BIC { get; set; }
        public string Bankleitzahl { get; set; }
        public string Bankname { get; set; }
        public string Kontoinhaber { get; set; }
    }
}
