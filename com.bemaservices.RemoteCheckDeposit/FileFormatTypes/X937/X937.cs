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
    [Description( "Processes a batch export for the X937 Format." )]
    [Export(typeof(FileFormatTypeComponent))]
    [ExportMetadata("ComponentName", "X937" )]
    class X937 : X937V2DSTU
    {
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
            int currencyTypeCheckId = Rock.Web.Cache.DefinedValueCache.Get( Rock.SystemGuid.DefinedValue.CURRENCY_TYPE_CHECK ).Id;
            if ( transactions.Any( t => t.FinancialPaymentDetail.CurrencyTypeValueId != currencyTypeCheckId ) )
            {
                errorMessages.Add( "One or more transactions is not of type 'Check'." );
                return null;
            }

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
