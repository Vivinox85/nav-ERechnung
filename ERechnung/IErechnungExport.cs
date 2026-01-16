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
        void FillInvoiceHeader(string invoiceNumber, string buyerReference, string orderNo, DateTime invoiceDate, string currencyCode, DateTime deliveryDate, string paymentTerms, DateTime paymentDueDate, string deliveryNoteNo);

        [DispId(5)]
        void AddSeller(string name, string street, string zipCode, string city, string country, string vatID, string taxNumber, string contact, string id, string email, string phone);

        [DispId(6)]
        void AddBuyer(string name, string street, string street2, string zipCode, string city, string country, string vatID, string contact, string organizationUnit, string email, string phone, string id, string orderReferenceDocument);

        [DispId(7)]
        void AddLineItem(string id, string name, string description, string customerID, string lineID, double quantity, string quantityCode, double unitPrice, double unitQuantity, string taxCategory, string taxType, double taxPercent, double lineTotal, string originCountry, double discountAmount = 0, double discountPercent = 0, double cuPreisDel = 0);

        [DispId(8)]
        void AddBankAccount(string iban, string bic, string bankleitzahl, string bankname, string kontoinhaber);

        [DispId(9)]
        void AddDeliveryAddress(string name, string street, string street2, string postcode, string city, string country);

        [DispId(10)]
        void AddSkonto(int dueDays, double skontoPercent);

        [DispId(11)]
        void AddInvoiceNote(string text, string subjectCode);

        [DispId(12)]
        void CreatePDF(string inPDFPath, string outPDFPath);

        [DispId(13)]
        void AddCharge(double actualAmount, string reason, string taxCategory, string taxType, double taxPercent, string reasonCode);

        [DispId(14)]
        void AddLineItemCharacteristic(string description, string value);
    }
}
