using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;

using Rock;
using Rock.Attribute;
using Rock.Model;

using com.bemaservices.RemoteCheckDeposit.Model;
using com.bemaservices.RemoteCheckDeposit.Records.X937;
using com.bemaservices.RemoteCheckDeposit;
using System.Text;

namespace com.bemaservices.RemoteCheckDeposit.FileFormatTypes
{

    /// <summary>
    /// Defines the x937 File export for Cass Commercial Bank
    /// </summary>
    [Description("Processes a batch export for Cass Commercial Bank.")]
    [Export(typeof(FileFormatTypeComponent))]
    [ExportMetadata("ComponentName", "Cass Commercial Bank")]
    [BooleanField(name: "Count the Deposit Ticket",
        description: "Set this to true to include the deposit slip in the bundle count.  Default is *usually* false.",
        defaultValue: false,
        key: "CountDepositSlip")]
    //[EncryptedTextField("Origin Routing Number", "Used on Type 10 Record 3 for Account Routing", true, key: "OriginRoutingNumber")]
    //[EncryptedTextField("Destination Routing Number", "", true, "072000096", key: "DestinationRoutingNumber")]
    public class CassCommercialBank : X937DSTU
    {

        #region Private Members

        /// <summary>
        /// Counts the deposit slip in the item counts
        /// </summary>
        private bool countDepositSlip = false;
        #endregion


        #region System Setting Keys
        /// <summary>
        /// The system setting for the next cash header identifier. These should never be
        /// repeated. Ever.
        /// </summary>
        protected const string SystemSettingNextCashHeaderId = "CassCommercial.NextCashHeaderId";

        /// <summary>
        /// The system setting that contains the last file modifier we used.
        /// </summary>
        protected const string SystemSettingLastFileModifier = "CassCommercial.LastFileModifier";

        /// <summary>
        /// The last item sequence number used for items.
        /// </summary>
        protected const string LastItemSequenceNumberKey = "CassCommercial.LastItemSequenceNumber";
        #endregion

        /// <summary>
        /// Gets the File Header Record (type 01)
        /// </summary>
        protected override FileHeader GetFileHeaderRecord(ExportOptions options)
        {
            var header = base.GetFileHeaderRecord(options);

            var originRoutingNumber = Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "InstitutionRoutingNumber"));

            // Override account number with institution routing number for Field 5
            header.ImmediateOriginRoutingNumber = originRoutingNumber;

            return header;
        }

        /// <summary>
        /// Gets the Cash Letter Header Record (type 10)
        /// </summary>
        protected override CashLetterHeader GetCashLetterHeaderRecord(ExportOptions options)
        {
            int cashHeaderId = GetSystemSetting(SystemSettingNextCashHeaderId).AsIntegerOrNull() ?? 0;
            var originRoutingNumber = Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "InstitutionRoutingNumber"));

            var header = base.GetCashLetterHeaderRecord(options);
            header.ID = cashHeaderId.ToString("D8");
            SetSystemSetting(SystemSettingNextCashHeaderId, (cashHeaderId + 1).ToString());

            // Override routing number with institution routing number for Field 4
            header.ClientInstitutionRoutingNumber = originRoutingNumber;

            return header;
        }

        /// <summary>
        /// Gets the Bundle Header Record (type 20)
        /// </summary>
        protected override BundleHeader GetBundleHeader(ExportOptions options, int bundleIndex)
        {
            var header = base.GetBundleHeader(options, bundleIndex);
            var originRoutingNumber = Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "InstitutionRoutingNumber"));

            // Override routing number with institution routing number for Field 4
            header.ClientInstitutionRoutingNumber = originRoutingNumber;

            // Set Bundle ID  (should be same as Bundle Sequence Number )
            header.ID = (bundleIndex + 1).ToString();

            return header;
        }

        /// <summary>
        /// Gets the item detail records (type 25)
        /// </summary>
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
            var detail = new Records.X937.CheckDetail
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
                BankOfFirstDepositIndicator = "Y",
                CheckDetailRecordAddendumCount = 00, // From Veronica 
                CorrectionIndicator = string.Empty,
                ArchiveTypeIndicator = string.Empty

            };

            return new List<Record> { detail };
        }
      

        /// <summary>
        /// Gets the image record for a specific transaction image (type 50 and 52).
        /// </summary>
        protected override List<Record> GetImageRecords(ExportOptions options, FinancialTransaction transaction, FinancialTransactionImage image, bool isFront)
        {
            var institutionRoutingNumber = Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "InstitutionRoutingNumber"));
            var records = base.GetImageRecords(options, transaction, image, isFront);

            foreach (var imageData in records.Where(r => r.RecordType == 52).Cast<dynamic>())
            {
                imageData.InstitutionRoutingNumber = institutionRoutingNumber;
            }

            foreach (var imageData in records.Where(r => r.RecordType == 50).Cast<dynamic>())
            {
                imageData.ImageCreatorRoutingNumber = institutionRoutingNumber;
            }

            return records;
        }

        /// <summary>
        /// Gets the credit detail deposit record (type 61).
        /// </summary>
        protected override List<Record> GetCreditDetailRecords(ExportOptions options, int bundleIndex, List<FinancialTransaction> transactions)
        {
            var accountNumber = Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "AccountNumber"));
            var routingNumber = Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "RoutingNumber"));
            var payorRoutingNumber = Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "PayorRoutingNumber"));
            var destinationRoutingNumber = Rock.Security.Encryption.DecryptString(GetAttributeValue(options.FileFormat, "DestinationRoutingNumber"));
            var records = new List<Record>();

            var creditDetail = new CreditDetail
            {
                PayorRoutingNumber = payorRoutingNumber,
                CreditAccountNumber = accountNumber + "/",
                Amount = transactions.Sum(t => t.TotalAmount),
                InstitutionItemSequenceNumber = GetNextItemSequenceNumber().ToString("000000000000000"),
                DocumentTypeIndicator = "G",
                SourceOfWorkCode = "3",
                DebitCreditIndicator = "2"
            };
            records.Add(creditDetail);

            for (int i = 0; i < 2; i++)
            {
                using (var ms = GetDepositSlipImage(options, creditDetail, i == 0))
                {
                    var tiffImageBytes = ConvertImageToTiffG4(ms).ReadBytesToEnd();
                    //
                    // Get the Image View Detail record (type 50).
                    //
                    var detail = new ImageViewDetail
                    {
                        ImageIndicator = 1,
                        ImageCreatorRoutingNumber = destinationRoutingNumber,
                        ImageCreatorDate = options.ExportDateTime,
                        ImageViewFormatIndicator = 0,
                        DataSize = (int)tiffImageBytes.Length,
                        CompressionAlgorithmIdentifier = 0,
                        SideIndicator = i,
                        ViewDescriptor = 0,
                        DigitalSignatureIndicator = 0
                    };

                    //
                    // Get the Image View Data record (type 52).
                    //
                    var data = new ImageViewData
                    {
                        InstitutionRoutingNumber = routingNumber,
                        CycleNumber = string.Empty,
                        BundleBusinessDate = options.BusinessDateTime,
                        ClientInstitutionItemSequenceNumber = creditDetail.InstitutionItemSequenceNumber,
                        ClippingOrigin = 0,
                        ImageData = ms.ReadBytesToEnd()
                    };

                    records.Add(detail);
                    records.Add(data);
                }
            }

            return records;
        }

        /// <summary>
        /// Gets the bundle control record (type 70).
        /// </summary>

        protected override Records.X937.BundleControl GetBundleControl(ExportOptions options, List<Record> records)
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
            var control = new Records.X937.BundleControl
            {
                ItemCount = itemRecords.Count(),
                TotalAmount = checkDetailRecords.Sum(r => (decimal)r.ItemAmount),
                MICRValidTotalAmount = 000000000000, //checkDetailRecords.Sum(r => (decimal)r.ItemAmount),
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
        protected override Records.X937.CashLetterControl GetCashLetterControlRecord(ExportOptions options, List<Record> records)
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
            var control = new Records.X937.CashLetterControl
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

        /// <summary>
        /// Gets the cash letter control record (type 99).
        /// </summary>
        protected override FileControl GetFileControlRecord(ExportOptions options, List<Record> records)
        {
            var fileControl = base.GetFileControlRecord(options, records);
            var itemRecords = records.Where(r => r.RecordType == 25); // Only count record types 25.
            // Some banks *might* include the deposit slip
            if (this.countDepositSlip)
            {
                itemRecords = records.Where(r => r.RecordType == 25 && r.RecordType == 61); // Just Overwrite.
            }

            fileControl.ImmediateOriginContactName = "".PadRight(14, ' ');
            fileControl.ImmediateOriginContactPhoneNumber = "0".PadRight(10, ' ');
            fileControl.TotalItemCount = itemRecords.Count();

            return fileControl;
        }

        #region Helper Methods
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
        /// <returns>A stream that contains the image data in TIFF 6.0 CCITT Group 4 format.</returns>
        protected virtual Stream GetDepositSlipImage(ExportOptions options, CreditDetail creditDetail, bool isFrontSide)
        {
            var bitmap = new System.Drawing.Bitmap(1200, 550);
            var g = System.Drawing.Graphics.FromImage(bitmap);

            var depositSlipTemplate = GetAttributeValue(options.FileFormat, "DepositSlipTemplate");
            var mergeFields = new Dictionary<string, object>
            {
                { "FileFormat", options.FileFormat },
                { "Amount", creditDetail.Amount.ToString( "C" ) }
            };
            var depositSlipText = depositSlipTemplate.ResolveMergeFields(mergeFields);

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
        #endregion


        /// <summary>
        /// Hashes the string with SHA256.
        /// </summary>
        protected string HashString(string contents)
        {
            byte[] byteContents = Encoding.Unicode.GetBytes(contents);

            var hash = new System.Security.Cryptography.SHA256CryptoServiceProvider().ComputeHash(byteContents);

            return string.Join("", hash.Select(b => b.ToString("x2")).ToArray());
        }
    }
}
