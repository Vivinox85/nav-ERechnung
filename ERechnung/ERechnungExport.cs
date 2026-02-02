using ERechnung.Models;
using s2industries.ZUGFeRD;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace ERechnung
{
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("ERechnungExport")]
    [ComVisible(true)]
    public class ERechnungExport : IErechnungExport
    {
        private XRechnung xRechnung;

        public ERechnungExport()
        {
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            Reset();
        }

        public void CreateXML(string filePath)
        {
            this.xRechnung.CreateXML(filePath);
        }

        public void CreatePDF(string inPDFPath, string outPDFPath)
        {
            this.xRechnung.CreatePDF(inPDFPath, outPDFPath);
        }

        public void FillInvoiceHeader(string invoiceNumber, string buyerReference, string orderNo, DateTime orderDate, DateTime invoiceDate, string currencyCode, DateTime deliveryDate, string paymentTerms, DateTime paymentDueDate, string deliveryNoteNo)
        {
            CurrencyCodes currency;
            Enum.TryParse(currencyCode, out currency);

            this.xRechnung.Type = InvoiceType.Invoice;

            this.xRechnung.InvoiceNumber = invoiceNumber;
            this.xRechnung.BuyerReference = buyerReference;
            this.xRechnung.InvoiceDate = invoiceDate;
            this.xRechnung.Currency = currency;
            this.xRechnung.DeliveryDate = deliveryDate;
            this.xRechnung.PaymentTerms = paymentTerms;
            this.xRechnung.PaymentDueDate = paymentDueDate;
            this.xRechnung.OrderNo = orderNo;
            this.xRechnung.OrderDate = orderDate;
            this.xRechnung.DeliveryNoteNo = deliveryNoteNo;
        }

        public void AddSeller(string name, string street, string zipCode, string city, string country, string vatID, string taxNumber, string contact, string id, string email, string phone, string vendorNo)
        {
            CountryCodes countryCode;
            Enum.TryParse(country, out countryCode);

            this.xRechnung.Seller = new Seller()
            {
                Name = name,
                Street = street,
                ZipCode = zipCode,
                City = city,
                Country = countryCode,
                VATID = vatID,
                TaxRegistrationSchemeID = TaxRegistrationSchemeID.VA,
                Contact = contact,
                ID = id,
                Email = email,
                Phone = phone,
                TaxNumber = taxNumber,
                VendorNo = vendorNo
            };
        }

        public void AddBuyer(string name, string street, string street2, string zipCode, string city, string country, string vatID, string contact, string organizationUnit, string email, string phone, string id, string orderReferenceDocument)
        {
            CountryCodes countryCode;
            Enum.TryParse(country, out countryCode);
            this.xRechnung.Buyer = new Buyer()
            {
                Name = name,
                Street = street,
                Street2 = street2,
                ZipCode = zipCode,
                City = city,
                Country = countryCode,
                VATID = vatID,
                TaxRegistrationSchemeID = TaxRegistrationSchemeID.VA,
                Contact = contact,
                OrganizationUnit = organizationUnit,
                Email = email,
                Phone = phone,
                ID = id,
                OrderReferenceDocument = orderReferenceDocument
            };
        }

        public void AddLineItem(string id, string name, string description, string customerID, string lineID, double quantity, string quantityCode, double unitPrice, double unitQuantity, string taxCategory, string taxType, double taxPercent, double lineTotal, string originCountry, double discountAmount=0, double discountPercent=0, double cuPreisDel=0)
        {
            QuantityCodes qc;
            TaxCategoryCodes tc;
            TaxTypes tt;
            CountryCodes originCountryCode;

            if(quantityCode == "PC")
            {
                quantityCode = "C62";
            }

            Enum.TryParse(quantityCode, out qc);
            Enum.TryParse(taxCategory, out tc);
            Enum.TryParse(taxType, out tt);
            Enum.TryParse(originCountry, out originCountryCode);

            this.xRechnung.LineItems.Add(new LineItem()
            {
                ID = id,
                Name = name,
                Description = description,
                CustomerID = customerID,
                LineID = lineID,
                Quantity = (decimal)quantity,
                Unit = qc,
                UnitPrice = (decimal)unitPrice,                
                TaxCategory = tc,
                TaxType = tt,
                TaxPercent = (decimal)taxPercent,
                LineTotal = (decimal)lineTotal,
                OriginCountry = originCountryCode,
                UnitQuantity = (decimal)unitQuantity,
                DiscountAmount = (decimal)discountAmount,
                DiscountPercent = (decimal)discountPercent,
                CUPreisDel = (decimal)cuPreisDel
            });
        }        

        public void AddBankAccount(string iban, string bic, string bankleitzahl, string bankname, string kontoinhaber)
        {
            this.xRechnung.BankAccounts.Add(new Bankkonto()
            {
                BIC = bic,
                IBAN = iban,
                Bankleitzahl = bankleitzahl,
                Bankname = bankname,
                Kontoinhaber = kontoinhaber
            });
        }

        public void AddDeliveryAddress(string name, string street, string street2, string postcode, string city, string country)
        {
            CountryCodes countryCode;
            Enum.TryParse(country, out countryCode);
            this.xRechnung.DeliveryAddress = new Party()
            {
                Name = name,
                ContactName = street,
                Street = street2,
                Postcode = postcode,
                City = city,
                Country = countryCode                
            };
        }

        public void AddSkonto(int dueDays, double skontoPercent)
        {
            this.xRechnung.SkontoOptions.Add(new PaymentTerms()
            {
                DueDays = dueDays,
                PaymentTermsType = PaymentTermsType.Skonto,
                Percentage = (decimal)skontoPercent
            });
        }

        public void Reset()
        {
            this.xRechnung = new XRechnung();
            this.xRechnung.LineItems = new List<LineItem>();
            this.xRechnung.BankAccounts = new List<Bankkonto>();
            this.xRechnung.SkontoOptions = new List<PaymentTerms>();
            this.xRechnung.Notes = new List<Note>();
            this.xRechnung.TradeCharges = new List<InvoiceCharge>();
            this.xRechnung.Taxes = new List<TaxApplication>();
        }

        public void AddInvoiceNote(string text, string subjectCode)
        {
            SubjectCodes subCode;
            Enum.TryParse(subjectCode, out subCode);
            this.xRechnung.Notes.Add(new Note(text, subjectCode: subCode));
        }

        public void AddCharge(double actualAmount, string reason, string taxCategory, string taxType, double taxPercent, string reasonCode)
        {
            TaxCategoryCodes tc;
            TaxTypes tt;
            ChargeReasonCodes rc;

            Enum.TryParse(taxCategory, out tc);
            Enum.TryParse(taxType, out tt);
            Enum.TryParse(reasonCode, out rc);

            this.xRechnung.TradeCharges.Add(new InvoiceCharge()
            {
                ActualAmount = (decimal)actualAmount,
                Reason = reason,
                TaxCategoryCode = tc,
                TaxTypeCode = tt,
                TaxPercent = (decimal)taxPercent,
                ChargeReasonCode = rc
            });
        }

        public void AddLineItemCharacteristic(string description, string value)
        {
            this.xRechnung.LineItems[this.xRechnung.LineItems.Count - 1].ItemCharacteristics.Add(new ItemCharacteristic()
            {
                Description = description,
                Value = value
            });
        }

        public void CreateZugferdXML(string filePath)
        {
            this.xRechnung.CreateZugferdXML(filePath);
        }

        public void CreateCreditMemoXML(string filePath)
        {
            this.xRechnung.CreateCreditMemoXML(filePath);
        }

        public void AddLineItemDetails(string orderNo, string orderLineNo, DateTime orderDate, string deliveryNo, string deliveryLineNo, DateTime deliveryDate)
        {
            int lastLineItemIndex = this.xRechnung.LineItems.Count - 1;
            this.xRechnung.LineItems[lastLineItemIndex].OrderNo = orderNo;
            this.xRechnung.LineItems[lastLineItemIndex].OrderLineID = orderLineNo;
            this.xRechnung.LineItems[lastLineItemIndex].OrderDate = orderDate;
            this.xRechnung.LineItems[lastLineItemIndex].DeliveryNo = deliveryNo;
            this.xRechnung.LineItems[lastLineItemIndex].DeliveryDate = deliveryDate;
            this.xRechnung.LineItems[lastLineItemIndex].DeliveryLineID = deliveryLineNo;
        }

        public void FillCreditMemoHeader(string cmNumber, string buyerReference, DateTime cmDate, string currencyCode)
        {
            CurrencyCodes currency;
            Enum.TryParse(currencyCode, out currency);

            this.xRechnung.Type = InvoiceType.CreditNote;

            this.xRechnung.InvoiceNumber = cmNumber;
            this.xRechnung.BuyerReference = buyerReference;
            this.xRechnung.InvoiceDate = cmDate;
            this.xRechnung.Currency = currency;            
        }

        public void AddLineItemNote(string text)
        {
            int lastLineItemIndex = this.xRechnung.LineItems.Count - 1;            
            this.xRechnung.LineItems[lastLineItemIndex].Notes.Add(new Note(text));
        }
        
        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            AssemblyName assemblyName = new AssemblyName(args.Name);
            string folderPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string assemblyPath = Path.Combine(folderPath, assemblyName.Name + ".dll");

            if (File.Exists(assemblyPath))
            {
                return Assembly.LoadFrom(assemblyPath);
            }

            return null;
        }
    }
}
