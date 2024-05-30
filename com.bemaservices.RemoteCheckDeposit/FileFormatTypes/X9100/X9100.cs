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
    [Description( "Processes a batch export for the X9100 Format.  This exports is built on X9.100-187-2008 " )]
    [Export(typeof(FileFormatTypeComponent))]
    [ExportMetadata("ComponentName", "X9100" )]
    class X9100 : X9100V2DSTU
    {
        #region Private Members

        const string BANK_NAME_KEY = "X9100";

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

        #endregion

        #region Bundle Records

        #endregion

        #region Private Methods

        /// <summary>
        /// Setup up the options
        /// </summary>
        /// <param name="options">ExportOptions</param>
        private void Setup(ExportOptions options)
        {
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
