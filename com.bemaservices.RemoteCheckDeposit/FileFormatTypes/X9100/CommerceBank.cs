using com.bemaservices.RemoteCheckDeposit;
using com.bemaservices.RemoteCheckDeposit.FileFormatTypes;
using com.bemaservices.RemoteCheckDeposit.Records.X9100;
using Rock;
using Rock.Attribute;
using Rock.Data;
using Rock.Model;
using Rock.Web.Cache;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace com.bemaservices.RemoteCheckDeposit.FileFormatTypes
{
    /// <summary>
    /// Defines the basic functionality of any component that will be exporting using the X9.100
    /// DSTU standard.
    /// </summary>
    [Description( "Processes a batch export for Commerce Bank.  This exports is built on X9.100-187-2008 " )]
    [Export(typeof(FileFormatTypeComponent))]
    [ExportMetadata("ComponentName", "Commerce Bank")]

    [EncryptedTextField(name: "Commerce Bank Client Id",
        description: "CB Client ID or Client Institution Routing Number. (Record Type 20 Field 4)",
        key: "CommerceBankClientId",
        required: true)]

    [EncryptedTextField(name: "Unique Customer Id",
        required: true,
        description: "Unique customer id that should be provided on summary sheet.(Record Type 01 Field 5 Right Justified)",
        key: "ImmediateOriginRoutingNumber")]

    [EncryptedTextField(name: "Immediate Origin Routing Number",
        description: "Immediate Origin Routing Number (Record Type 01 Field 5)",
        required: true,
        key: "ImmediateOriginRoutingNumber")]

    [EncryptedTextField(name: "Routing Transit Number",
        required: false,
        description: "The routing transit number or 'RT' (Record Type 61 Field 5)",
        key: "RoutingTransitNumber")]

    [EncryptedTextField(name: "Posting Account Number", 
        description: "", 
        key: "postingAccountNumber")]

    [TextField(name: "Collection Type Indicator",
        required: true,
        description: "Cash Letter Header Record (Type 10) Field value must be '00-06', '12', or '90'",
        key: "CollectionTypeValue")]

    [BooleanField(name: "Count the Deposit Ticket",
        description: "Set this to true to include the deposit slip in the bundle count.  Default is *usually* false.",
        defaultValue: false,
        key: "CountDepositSlip")]

    [CodeEditorField("Deposit Slip Template", "The template for the deposit slip that will be generated. <span class='tip tip-lava'></span>",
        Rock.Web.UI.Controls.CodeEditorMode.Lava,
        defaultValue: @"Customer: {{ FileFormat | Attribute:'OriginName' }}
Account: {{ FileFormat | Attribute:'AccountNumber' }}
Amount: {{ Amount }}", order: 30)]
    class CommerceBank : X9100DSTU
    {
        #region Private Members

        const string BANK_NAME_KEY = "CommerceBank";

        /// <summary>
        /// Account Number
        /// </summary>
        private string accountNumber = string.Empty;

        /// <summary>
        /// CB Client Id Record Type 20 Field 4
        /// </summary>
        private string commerceBankClientId = string.Empty;

        /// <summary>
        /// Routing Number
        /// </summary>
        private string routingNumber = string.Empty;

        /// <summary>
        /// The Routing Transit number Rec
        /// </summary>
        private int routingTransitNumber = 0;

        /// <summary>
        /// The immediate origin routing number Record Type Field 5
        /// </summary>
        private string originRoutingNumber = string.Empty;

        /// <summary>
        /// Collection Type Indicator
        /// </summary>
        private int collectionTypeIndicator = 00;

        /// <summary>
        /// Counts the deposit slip in the item counts
        /// </summary>
        private bool countDepositSlip = false;

        /// <summary>
        /// Contact Name
        /// </summary>
        private string contactName = string.Empty;

        /// <summary>
        /// Contact Phone
        /// </summary>
        private string contactPhone = string.Empty;

        /// <summary>
        /// Location Field used for identifying different stores or accounts on CB.  
        /// </summary>
        private string locationField = string.Empty;

        /// <summary>
        /// Arbitrary cycle number that must match in different record types.
        /// </summary>
        private string cycleNumber = string.Empty;

        /// <summary>
        /// Field Header Record (Type 01) Field 5 unique customer id right justified
        /// </summary>
        private string uniqueCustomerId = string.Empty;

        #endregion

        #region System Setting Keys

        /// <summary>
        /// The system setting for the next cash header identifier. These should never be
        /// repeated. Ever.
        /// </summary>
        protected const string SystemSettingNextCashHeaderId = BANK_NAME_KEY + ".NextCashHeaderId";

        /// <summary>
        /// The system setting that contains the last file modifier we used.
        /// </summary>
        protected const string SystemSettingLastFileModifier = BANK_NAME_KEY + ".LastFileModifier";

        /// <summary>
        /// The last item sequence number used for items.
        /// </summary>
        protected const string LastItemSequenceNumberKey = BANK_NAME_KEY + ".LastItemSequenceNumber";

        #endregion

        #region Export Batches

        /// <summary>
        /// Exports a collection of batches to a binary file that can be downloaded by the user
        /// and sent to their financial institution. The returned BinaryFile should not have been
        /// saved to the database yet.
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="errorMessages">On return will contain a list of error messages if not empty.</param>
        /// <returns>
        /// A <see cref="Stream" /> of data that should be downloaded to the user in a file.
        /// </returns>
        public override Stream ExportBatches(ExportOptions options, out List<string> errorMessages)
        {
            Setup(options);

            var records = new List<Record>();

            errorMessages = new List<string>();

            //
            // Get all the transactions that will be exported from these batches.
            //
            var transactions = options.Batches.SelectMany(b => b.Transactions)
                .OrderBy(t => t.ProcessedDateTime)
                .ThenBy(t => t.Id)
                .ToList();

            //
            // Perform error checking to ensure that all the transactions in these batches
            // are of the proper currency type.
            //
            //int currencyTypeCheckId = Rock.Web.Cache.DefinedValueCache.Get(Rock.SystemGuid.DefinedValue.CURRENCY_TYPE_CHECK).Id;
            //if (transactions.Any(t => t.FinancialPaymentDetail.CurrencyTypeValueId != currencyTypeCheckId))
            //{
            //    errorMessages.Add("One or more transactions is not of type 'Check'.");
            //    return null;
            //}

            //
            // Generate all the X9.100 records for this set of transactions.
            //

            //
            // Record Type 01
            //
            records.Add(GetFileHeaderRecord(options));

            //
            // Record Type 10
            //
            records.Add(GetCashLetterHeaderRecord(options));

            //
            // Record Type 20
            //
            records.AddRange(GetBundleRecords(options, transactions));

            //
            // Record Type 70
            //
            records.Add(GetCashLetterControlRecord(options, records));

            //
            // Record Type 90
            //
            records.Add(GetFileControlRecord(options, records));


            return GetDataStream(records);
        }


        #endregion

        #region File Records

        /// <summary>
        /// Gets the bundle header record (type 20).
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="bundleIndex">Number of existing bundle records in the cash letter.</param>
        /// <returns>A BundleHeader record.</returns>
        protected override BundleHeader GetBundleHeader(ExportOptions options, int bundleIndex)
        {
            //var routingNumber = Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "RoutingNumber"));

            var header = new BundleHeader
            {
                CollectionTypeIndicator = this.collectionTypeIndicator,
                DestinationRoutingNumber = this.routingNumber,
                ClientInstitutionRoutingNumber = this.commerceBankClientId,
                BusinessDate = options.BusinessDateTime,
                CreationDate = options.ExportDateTime,
                ID = bundleIndex.ToString(),
                SequenceNumber = (bundleIndex + 1).ToString(),
                CycleNumber = this.cycleNumber,
                ReturnLocationRoutingNumber = string.Empty, 
                UserField = string.Empty,
            };

            return header;
        }

        /// <summary>
        /// Gets the item detail records (type 25)
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="transaction">The transaction being deposited.</param>
        /// <returns>A collection of records.</returns>
        protected override List<Record> GetItemDetailRecords(ExportOptions options, FinancialTransaction transaction)
        {
            //var accountNumber = Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "AccountNumber"));
            //var routingNumber = Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "RoutingNumber"));

            //
            // Parse the MICR data from the transaction.
            //
            var micr = GetMicrInstance(transaction.CheckMicrEncrypted);

            var transactionRoutingNumber = micr.GetRoutingNumber();

            // Check Detail Record Type (25)
            var detail = new CheckDetail
            {
                AuxiliaryOnUs = micr.GetAuxOnUs(),
                ExternalProcessingCode = micr.GetExternalProcessingCode(),
                PayorBankRoutingNumber = transactionRoutingNumber.Substring(0, 8),
                PayorBankRoutingNumberCheckDigit = transactionRoutingNumber.Substring(8, 1),
                OnUs = string.Format("{0}/{1}", micr.GetAccountNumber(), micr.GetCheckNumber()),
                ItemAmount = transaction.TotalAmount,
                ClientInstitutionItemSequenceNumber = transaction.Id.ToString(),
                DocumentationTypeIndicator = "G", // Field Value must be "G" - Meaning there are 2 images present
                ElectronicReturnAcceptanceIndicator = string.Empty,
                MICRValidIndicator = null,
                BankOfFirstDepositIndicator = "U",
                CheckDetailRecordAddendumCount = 00, // From Veronica 
                CorrectionIndicator = string.Empty,
                ArchiveTypeIndicator = string.Empty

            };

            return new List<Record> { detail };
        }

        /// <summary>
        /// Gets the bundle control record (type 70).
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="records">The existing records in the bundle.</param>
        /// <returns>A BundleControl record.</returns>
        protected override BundleControl GetBundleControl(ExportOptions options, List<Record> records)
        {
            var itemRecords = records.Where(r => r.RecordType == 25);

            // If we are including the credit items then we need to count those as well.
            if (this.countDepositSlip)
            {
                itemRecords = records.Where(r => r.RecordType == 25 && r.RecordType == 61);  // Just Overwrite
            }

            var checkDetailRecords = records.Where(r => r.RecordType == 25).Cast<dynamic>(); // Only count checks and not the credit detail
            var imageDetailRecords = records.Where(r => r.RecordType == 52);

            // Record Type 70
            var control = new BundleControl
            {
                ItemCount = itemRecords.Count(),
                TotalAmount = checkDetailRecords.Sum(r => (decimal)r.ItemAmount),
                MICRValidTotalAmount = 000000000000,
                ImageCount = imageDetailRecords.Count()
            };

            return control;
        }

        /// <summary>
        /// Gets the cash letter control record (type 90).
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="records">Existing records in the cash letter.</param>
        /// <returns>A CashLetterControl record.</returns>
        protected override CashLetterControl GetCashLetterControlRecord(ExportOptions options, List<Record> records)
        {
            var bundleHeaderRecords = records.Where(r => r.RecordType == 20);
            var checkDetailRecords = records.Where(r => r.RecordType == 25).Cast<dynamic>();
            var itemRecords = records.Where(r => r.RecordType == 25); // Only count record types 25.
            var imageDetailRecords = records.Where(r => r.RecordType == 52);
            var organizationName = GetAttributeValue(options.FileFormat, "OriginName");

            // Some banks *might* include the deposit slip
            if (this.countDepositSlip)
            {
                itemRecords = records.Where(r => r.RecordType == 25 && r.RecordType == 61); // Just Overwrite.
            }

            // Record Type 90
            var control = new CashLetterControl
            {
                BundleCount = bundleHeaderRecords.Count(),
                ItemCount = itemRecords.Count(),
                //TotalAmount = 000000000000,//checkDetailRecords.Sum(c => (decimal)c.ItemAmount), // Must be 0's  According to Veronica, This should now be the total amount
                TotalAmount = checkDetailRecords.Sum(c => (decimal)c.ItemAmount), // According to Veronica 09-28-2018
                ImageCount = imageDetailRecords.Count(),
                ECEInstitutionName = organizationName
            };

            return control;
        }

        #endregion

        #region Methods

        /// <summary>
        /// Gets the next item sequence number.
        /// </summary>
        /// <returns>An integer that identifies the unique item sequence number that can be used.</returns>
        protected int GetNextItemSequenceNumber()
        {
            int lastSequence = GetSystemSetting(LastItemSequenceNumberKey).AsIntegerOrNull() ?? 0;
            int nextSequence = lastSequence + 1;

            SetSystemSetting(LastItemSequenceNumberKey, nextSequence.ToString());

            return nextSequence;
        }

        /// <summary>
        /// Gets the credit detail deposit record (type 61).
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="transactions">The transactions associated with this deposit.</param>
        /// <param name="isFrontSide">True if the image to be retrieved is the front image.</param>
        /// <returns>A stream that contains the image data in TIFF 6.0 CCITT Group 4 format.</returns>
        protected virtual Stream GetDepositSlipImage(ExportOptions options, CreditReconciliation creditDetail, bool isFrontSide)
        {
            var bitmap = new System.Drawing.Bitmap(1200, 550);
            var g = System.Drawing.Graphics.FromImage(bitmap);

            var depositSlipTemplate = GetAttributeValue(options.FileFormat, "DepositSlipTemplate");
            var mergeFields = new Dictionary<string, object>
            {
                { "FileFormat", options.FileFormat },
                { "Amount", creditDetail.ItemAmount.ToString( "C" ) }
            };
            var depositSlipText = depositSlipTemplate.ResolveMergeFields(mergeFields, null);

            //
            // Ensure we are opague with white.
            //
            g.FillRectangle(System.Drawing.Brushes.White, new System.Drawing.Rectangle(0, 0, 1200, 550));

            if (isFrontSide)
            {
                g.DrawString(depositSlipText,
                    new System.Drawing.Font("Tahoma", 30),
                    System.Drawing.Brushes.Black,
                    new System.Drawing.PointF(50, 50));
            }

            g.Flush();

            //
            // Ensure the DPI is correct.
            //
            bitmap.SetResolution(200, 200);

            //
            // Compress using TIFF, CCITT Group 4 format.
            //
            var codecInfo = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                .Where(c => c.MimeType == "image/tiff")
                .First();
            var parameters = new System.Drawing.Imaging.EncoderParameters(1);
            parameters.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)System.Drawing.Imaging.EncoderValue.CompressionCCITT4);

            var ms = new MemoryStream();
            bitmap.Save(ms, codecInfo, parameters);
            ms.Position = 0;

            return ms;
        }

        /// <summary>
        /// Gets the records that identify a single check being deposited.
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="transaction">The transaction to be deposited.</param>
        /// <returns>
        /// A collection of records.
        /// </returns>
        protected override List<Record> GetItemRecords(ExportOptions options, FinancialTransaction transaction)
        {
            var records = new List<Record>();

            // Record Type 25
            records.AddRange(GetItemDetailRecords(options, transaction));

            records.AddRange(GetImageRecords(options, transaction, transaction.Images.Take(1).First(), true));
            records.AddRange(GetImageRecords(options, transaction, transaction.Images.Skip(1).Take(1).First(), false));

            var sequenceNumber = GetNextItemSequenceNumber();

            //
            // Modify the Check Detail Record and Check Image Data records to have
            // a unique item sequence number.
            //
            var checkDetail = records.Where(r => r.RecordType == 25).Cast<dynamic>().FirstOrDefault();
            checkDetail.ClientInstitutionItemSequenceNumber = sequenceNumber.ToString("000000000000000");

            foreach (var imageData in records.Where(r => r.RecordType == 52).Cast<dynamic>())
            {
                imageData.ClientInstitutionItemSequenceNumber = sequenceNumber.ToString("000000000000000");
            }

            return records;
        }

        #endregion

        #region Bundle Records

        /// <summary>
        /// Gets all the bundle records in required for the transactions specified.
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="transactions">The transactions to be exported.</param>
        /// <returns>A collection of records that identify all the exported transactions.</returns>
        protected override List<Record> GetBundleRecords(ExportOptions options, List<FinancialTransaction> transactions)
        {
            var records = new List<Record>();

            for (int bundleIndex = 0; (bundleIndex * MaxItemsPerBundle) < transactions.Count(); bundleIndex++)
            {
                var bundleRecords = new List<Record>();
                var bundleTransactions = transactions.Skip(bundleIndex * MaxItemsPerBundle)
                    .Take(MaxItemsPerBundle)
                    .ToList();

                //
                // Add the bundle header for this set of transactions. Record Type 20
                //
                bundleRecords.Add(GetBundleHeader(options, bundleIndex));

                //
                // Add records for each transaction in the bundle.
                //
                foreach (var transaction in bundleTransactions)
                {
                    try
                    {
                        bundleRecords.AddRange(GetItemRecords(options, transaction));
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(string.Format("Error processing transaction {0}.", transaction.Id), ex);
                    }
                }

                //
                // Add the bundle control record.  Record Type 70
                //
                bundleRecords.Add(GetBundleControl(options, bundleRecords));

                records.AddRange(bundleRecords);
            }

            return records;
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Setup up the options
        /// </summary>
        /// <param name="options">ExportOptions</param>
        private void Setup(ExportOptions options)
        {
            this.cycleNumber = (((int)DateTime.Now.DayOfWeek == 0) ? 7 : (int)DateTime.Now.DayOfWeek).ToString();

            this.accountNumber = Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "AccountNumber"));
            this.routingNumber = Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "RoutingNumber"));
            this.originRoutingNumber = Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "ImmediateOriginRoutingNumber"));
            this.commerceBankClientId = Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "CommerceBankClientId"));
            this.contactName = GetAttributeValue(options.FileFormat, "ContactName");
            this.contactPhone = GetAttributeValue(options.FileFormat, "ContactPhone").Replace(" ", string.Empty);
            this.countDepositSlip = bool.Parse(GetAttributeValue(options.FileFormat, "CountDepositSlip"));

            // Try and get the routing transit number Record Type 61 Field 5
            int.TryParse(Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "RoutingTransitNumber")), 
                out this.routingTransitNumber);

            // Try and get an int value from the collection type value
            int.TryParse(GetAttributeValue(options.FileFormat, "CollectionTypeValue"), out this.collectionTypeIndicator);

            // Get the location Field from the dictionary 
            if (options.ExtraOptions.ContainsKey("LocationField"))
            {
                this.locationField = options.ExtraOptions["LocationField"].ToStringSafe().Trim();
            }
        }

        /// <summary>
        /// Gets the data stream
        /// </summary>
        /// <param name="records"></param>
        /// <returns></returns>
        private static Stream GetDataStream(List<Record> records)
        {
            //
            // Encode all the records into a memory stream so that it can be saved to a file
            // by the caller.
            //
            var stream = new MemoryStream();
            using (var writer = new BinaryWriter(stream, System.Text.Encoding.UTF8, true))
            {
                foreach (var record in records)
                {
                    record.Encode(writer);
                }
            }

            stream.Position = 0;

            return stream;
        }

        #endregion
    }
}
