using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ERechnung
{
    [Guid("A80254EA-429B-4BEE-B40F-9C9CB85D99DF")]
    [ComVisible(true)]
    public interface IErechnungExport
    {
        [DispId(2)]
        void Reset();

        [DispId(3)]
        void CreateXML(string filePath);

        [DispId(4)]
        void FillInvoiceHeader(string invoiceNumber, string orderNumber, DateTime invoiceDate, string currencyCode, DateTime deliveryDate, string paymentTerms, DateTime paymentDueDate);

        [DispId(5)]
        void AddSeller(string name, string street, string zipCode, string city, string country, string vatID, string taxNumber, string legalDescription, string contact, string id, string email, string phone);

        [DispId(6)]
        void AddBuyer(string name, string street, string zipCode, string city, string country, string vatID, string contact, string organizationUnit, string email, string phone, string id, string orderReferenceDocument);

        [DispId(7)]
        void AddLineItem(string id, string name, string description, string customerID, double quantity, string quantityCode, double unitPrice, string taxCategory, string taxType, double taxPercent, double lineTotal);

        [DispId(8)]
        void AddBankAccount(string iban, string bic, string bankleitzahl, string bankname, string kontoinhaber);

        [DispId(9)]
        void AddDeliveryAddress(string name, string street, string postcode, string city, string country);
    }
}
