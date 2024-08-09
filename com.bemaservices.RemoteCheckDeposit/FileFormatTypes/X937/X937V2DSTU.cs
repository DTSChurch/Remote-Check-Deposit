using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;

using Rock;
using Rock.Attribute;
using Rock.Model;
using Rock.Web.UI.Controls;

using com.bemaservices.RemoteCheckDeposit.Model;
using System.Drawing;
using System.Windows.Media.Imaging;
using com.bemaservices.RemoteCheckDeposit.Records.X937;
using System.Data.Entity.Validation;

namespace com.bemaservices.RemoteCheckDeposit.FileFormatTypes
{
    /// <summary>
    /// Defines the basic functionality of any component that will be exporting using the X9.37
    /// DSTU standard.
    /// </summary>    

    // Origin Settings
    [TextField( "Origin Name",
        Description = "The name of the church.",
        Key = AttributeKey.OriginName,
        IsRequired = false,
        DefaultValue = "",
        Order = 0,
        Category = "Origin Fields" )]
    [TextField( "Origin Contact Name",
        Description = "The name of the person the bank will contact if there are issues.",
        Key = AttributeKey.OriginContactName,
        IsRequired = false,
        DefaultValue = "",
        Order = 1,
        Category = "Origin Fields" )]
    [TextField( "Origin Contact Phone",
        Description = "The phone number the bank will call if there are issues.",
        Key = AttributeKey.OriginContactPhone,
        IsRequired = false,
        DefaultValue = "",
        Order = 2,
        Category = "Origin Fields" )]
    [EncryptedTextField( "Origin Routing Number",
        Description = "Your origin routing number.",
        Key = AttributeKey.OriginRoutingNumber,
        IsRequired = false,
        DefaultValue = "",
        Order = 3,
        Category = "Origin Fields" )]

    // Destination Settings
    [TextField( "Destination Name",
        Description = "The name of the bank the deposit will be made to.",
        Key = AttributeKey.DestinationName,
        IsRequired = false,
        DefaultValue = "",
        Order = 0,
        Category = "Destination Fields" )]
    [EncryptedTextField( "Destination Routing Number",
        Description = "The destination routing number.",
        Key = AttributeKey.DestinationRoutingNumber,
        IsRequired = false,
        DefaultValue = "",
        Order = 1,
        Category = "Destination Fields" )]

    // ECE Institution Settings
    [EncryptedTextField( "Institution Name",
        Description = "This is defined by your bank, it is typically but not always the same as the origin name.",
        Key = AttributeKey.InstitutionName,
        IsRequired = false,
        DefaultValue = "",
        Order = 0,
        Category = "ECE Institution Fields" )]
    [EncryptedTextField( "ECE Institution Routing Number",
        Description = "This is defined by your bank, it is typically but not always the same as the origin routing number",
        Key = AttributeKey.InstitutionRoutingNumber,
        IsRequired = false,
        DefaultValue = "",
        Order = 1,
        Category = "ECE Institution Fields" )]
    [CustomRadioListField( "Item Sequence Number Justification",
        Description = "Whether the Item Sequence Number should be Right or Left Justified. The default for most banks is right-justified.",
        Key = AttributeKey.ItemSequenceNumberJustification,
        ListSource = "Right,Left",
        DefaultValue = "Right",
        Order = 2,
        IsRequired = true,
        Category = "ECE Institution Fields" )]

    // BOFD Settings
    [EncryptedTextField( "BOFD Routing Number",
        Description = "This is defined by your bank, it is typically but not always the same as the ECE Institution routing number",
        Key = AttributeKey.BOFDRoutingNumber,
        IsRequired = false,
        DefaultValue = "",
        Order = 0,
        Category = "Bank of First Deposit Fields" )]
    [TextField( "Truncation Indicator",
        Description = "The Default for this value is 'N', but some banks require 'Y'.",
        Key = AttributeKey.TruncationIndicator,
        IsRequired = false,
        DefaultValue = "N",
        Order = 1,
        Category = "Bank of First Deposit Fields" )]

    // Credit Deposit Settings
    [EnumField( "Credit Record Type",
        Description = "What type of credit detail deposit record should be included if any. Type 61.A is an alternate version of Type 61 with a RecordUsageIndicator field.",
        Key = AttributeKey.CreditRecordType,
        IsRequired = true,
        Order = 0,
        EnumSourceType = typeof( CreditDetailRecordType ),
        DefaultEnumValue = ( int ) CreditDetailRecordType.None,
        Category = "Credit Deposit Settings" )]
    [IntegerField( "Credit Deposit Check Number",
        Description = "The check number to be appended onto the end of File Type 61 Field 5: On-Us. The Default is 20.",
        Key = AttributeKey.CreditDepositCheckNumber,
        IsRequired = false,
        DefaultValue = "20",
        Order = 1,
        Category = "Credit Deposit Settings" )]
    [CodeEditorField( "Deposit Slip Template",
        Description = "The template for the deposit slip that will be generated. <span class='tip tip-lava'></span>",
        Key = AttributeKey.DepositSlipTemplate,
        EditorMode = CodeEditorMode.Lava,
        Order = 2,
        IsRequired = false,
        Category = "Credit Deposit Settings",
        DefaultValue = @"Customer: {{ FileFormat | Attribute:'OriginName' }}
CICL-{{ FileFormat | Attribute:'AccountNumber' }}
Account: {{ FileFormat | Attribute:'AccountNumber' }}
Amount: {{ Amount }}
ItemCount: {{ ItemCount }}" )]

    // Rock Settings
    [BooleanField( "Test Mode",
        Description = "If true then the generated files will be marked as test-mode.",
        Key = AttributeKey.TestMode,
        IsRequired = true,
        Order = 0,
        Category = "Rock Settings" )]
    [DefinedValueField( "Currency Types To Export",
        DefinedTypeGuid = "1D1304DE-E83A-44AF-B11D-0C66DD600B81",
        Description = "Select which check types are valid to send in batches to export (Defaults to Checks)",
        Key = AttributeKey.CurrencyTypes,
        AllowMultiple = true,
        IsRequired = true,
        DefaultValue = "8B086A19-405A-451F-8D44-174E92D6B402",
        Order = 1,
        Category = "Rock Settings" )]
    [BooleanField( "Enable Digital Endorsement",
        Description = "Prints text on the back of the check digitally, as an endoresment of the check",
        Key = AttributeKey.EnableDigitalEndorsement,
        DefaultBooleanValue = false,
        Order = 2,
        Category = "Rock Settings" )]
    [CodeEditorField( "Check Endorsement Template",
        Description = "The template for the back of check endorsement. <span class='tip tip-lava'></span>",
        Key = AttributeKey.CheckEndorsementTemplate,
        EditorMode = CodeEditorMode.Lava,
        Order = 3,
        IsRequired = false,
        Category = "Rock Settings",
        DefaultValue = @"{{ FileFormat | Attribute:'OriginName' }}
Account: {{ FileFormat | Attribute:'AccountNumber' }}
Date: {{ BusinessDate | Date:'M/d/yyyy' }}" )]
    public abstract class X937V2DSTU : FileFormatTypeComponent
    {
        #region Attribute Keys
        private static class AttributeKey
        {
            // Origin Settings
            public const string OriginName = "OriginName";
            public const string OriginContactName = "OriginContactName";
            public const string OriginContactPhone = "OriginContactPhone";
            public const string OriginRoutingNumber = "OriginRoutingNumber";

            // Destination Settings
            public const string DestinationName = "DestinationName";
            public const string DestinationRoutingNumber = "DestinationRoutingNumber";

            // ECE Institution Settings
            public const string InstitutionName = "InstitutionName";
            public const string InstitutionRoutingNumber = "InstitutionRoutingNumber";
            public const string ItemSequenceNumberJustification = "ItemSequenceNumberJustification";

            // Bank of First Deposit Settings
            public const string BOFDRoutingNumber = "BOFDRoutingNumber";
            public const string TruncationIndicator = "TruncationIndicator";

            // Credit Deposit Settings
            public const string CreditRecordType = "CreditRecordType";
            public const string CreditDepositCheckNumber = "CreditDepositCheckNumber";
            public const string DepositSlipTemplate = "DepositSlipTemplate";

            // Rock Settings
            public const string TestMode = "TestMode";
            public const string CurrencyTypes = "CurrencyTypes";
            public const string EnableDigitalEndorsement = "EnableDigitalEndorsement";
            public const string CheckEndorsementTemplate = "CheckEndorsementTemplate";

            // Obsolete Keys
            public const string ObsoleteRoutingNumber = "RoutingNumber";
            public const string ObsoleteAccountNumber = "AccountNumber";
        }
        #endregion

        #region System Setting Keys

        /// <summary>
        /// The system setting for the next cash header identifier. These should never be
        /// repeated. Ever.
        /// </summary>
        protected const string SystemSettingNextCashHeaderId = "X937V2DSTU.NextCashHeaderId";

        /// <summary>
        /// The system setting that contains the last file modifier we used.
        /// </summary>
        protected const string SystemSettingLastFileModifier = "X937V2DSTU.LastFileModifier";
        /// <summary>
        /// The last item sequence number used for items.
        /// </summary>
        protected const string LastItemSequenceNumberKey = "X937V2DSTU.LastItemSequenceNumber";

        #endregion

        /// <summary>
        /// Gets the next item sequence number.
        /// </summary>
        /// <returns>An integer that identifies the unique item sequence number that can be used.</returns>
        protected int GetNextItemSequenceNumber()
        {
            int lastSequence = GetSystemSetting( LastItemSequenceNumberKey ).AsIntegerOrNull() ?? 0;
            int nextSequence = lastSequence + 1;

            SetSystemSetting( LastItemSequenceNumberKey, nextSequence.ToString() );

            return nextSequence;
        }

        /// <summary>
        /// Gets the maximum items per bundle. Most banks limit the number of checks that
        /// can exist in each bundle. This specifies what that maximum is.
        /// </summary>
        public virtual int MaxItemsPerBundle
        {
            get
            {
                return 200;
            }
        }

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
        public override Stream ExportBatches( ExportOptions options, out List<string> errorMessages )
        {
            var records = new List<Record>();

            errorMessages = new List<string>();

            //
            // Get all the transactions that will be exported from these batches.
            //
            var transactions = options.Batches.SelectMany( b => b.Transactions )
                .OrderBy( t => t.ProcessedDateTime )
                .ThenBy( t => t.Id )
                .ToList();

            //
            // Perform error checking to ensure that all the transactions in these batches
            // are of the proper currency type.
            //
            List<Guid> currencyGuids = GetAttributeValue( options.FileFormat, AttributeKey.CurrencyTypes ).SplitDelimitedValues().AsGuidList();
            if ( !currencyGuids.Any() )
            {
                //Add the default check option if nothing is selected
                currencyGuids.Add( Guid.Parse( "8B086A19-405A-451F-8D44-174E92D6B402" ) );
            }
            List<int> currencyIds = new List<int>();
            foreach ( Guid guid in currencyGuids )
            {
                currencyIds.Add( Rock.Web.Cache.DefinedValueCache.Get( guid ).Id );
            }

            if ( transactions.Any( t => !currencyIds.Contains( t.FinancialPaymentDetail.CurrencyTypeValueId ?? -1 ) ) )
            {
                errorMessages.Add( "One or more transactions is not of a selected Check type." );
                throw new Exception( "One or more transactions is not of a selected Check type." );
            }

            //
            // Generate all the X9.37 records for this set of transactions.
            //
            records.Add( GetFileHeaderRecord( options ) );
            records.Add( GetCashLetterHeaderRecord( options ) );
            records.AddRange( GetBundleRecords( options, transactions ) );
            records.Add( GetCashLetterControlRecord( options, records ) );
            records.Add( GetFileControlRecord( options, records ) );

            //
            // Encode all the records into a memory stream so that it can be saved to a file
            // by the caller.
            //
            var stream = new MemoryStream();
            using ( var writer = new BinaryWriter( stream, System.Text.Encoding.UTF8, true ) )
            {
                foreach ( var record in records )
                {
                    record.Encode( writer );
                }
            }

            stream.Position = 0;

            return stream;
        }

        #region File Records

        /// <summary>
        /// Gets the file header record (type 01).
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <returns>A FileHeader record.</returns>
        protected virtual Records.X937.FileHeader GetFileHeaderRecord( ExportOptions options )
        {
            string destinationRoutingNumber = GetValueWithFallback( options, AttributeKey.DestinationRoutingNumber, AttributeKey.ObsoleteRoutingNumber );
            string originRoutingNumber = GetValueWithFallback( options, AttributeKey.OriginRoutingNumber, AttributeKey.ObsoleteAccountNumber );

            var header = new Records.X937.FileHeader
            {
                StandardLevel = 03,
                FileTypeIndicator = GetAttributeValue( options.FileFormat, AttributeKey.TestMode ).AsBoolean( true ) ? "T" : "P",
                ImmediateDestinationRoutingNumber = destinationRoutingNumber,
                ImmediateOriginRoutingNumber = originRoutingNumber,
                FileCreationDateTime = options.ExportDateTime,
                ResendIndicator = "N",
                ImmediateDestinationName = GetAttributeValue( options.FileFormat, AttributeKey.DestinationName ),
                ImmediateOriginName = GetAttributeValue( options.FileFormat, AttributeKey.OriginName ),
                FileIdModifier = "1", // TODO: Need some way to track this and reset each day.
                CountryCode = "US", /* Should be safe, X9.37 is only used in the US as far as I know. */
                UserField = string.Empty
            };

            return header;
        }

        /// <summary>
        /// Gets the file control record (type 99).
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="records">The existing records in the file.</param>
        /// <returns>A FileControl record.</returns>
        protected virtual Records.X937.FileControl GetFileControlRecord( ExportOptions options, List<Record> records )
        {
            var cashHeaderRecords = records.Where( r => r.RecordType == 10 );
            var detailRecords = records.Where( r => r.RecordType == 25 ).Cast<dynamic>();
            //count deposit records as well as items
            //var itemRecords = records.Where( r => r.RecordType == 25 ); // Only count checks!
            var itemRecords = records.Where( r => r.RecordType == 25 || r.RecordType == 61 );

            var control = new Records.X937.FileControl
            {
                CashLetterCount = cashHeaderRecords.Count(),
                TotalRecordCount = records.Count + 1, /* Plus one to include self */
                TotalItemCount = itemRecords.Count(),
                TotalAmount = detailRecords.Sum( c => ( decimal ) c.ItemAmount ),
                ImmediateOriginContactName = GetAttributeValue( options.FileFormat, AttributeKey.OriginContactName ),
                ImmediateOriginContactPhoneNumber = GetAttributeValue( options.FileFormat, AttributeKey.OriginContactPhone ).Replace( " ", string.Empty )
            };

            return control;
        }

        #endregion

        #region Cash Letter Records

        /// <summary>
        /// Gets the cash letter header record (type 10).
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <returns>A CashLetterHeader record.</returns>
        /// 
        protected virtual Records.X937.CashLetterHeader GetCashLetterHeaderRecord( ExportOptions options )
        {
            var destinationRoutingNumber = int.Parse( GetValueWithFallback( options, AttributeKey.DestinationRoutingNumber, AttributeKey.ObsoleteRoutingNumber ) );
            var institutionRoutingNumber = int.Parse( GetValueWithFallback( options, AttributeKey.InstitutionRoutingNumber, AttributeKey.ObsoleteAccountNumber ) );
            var contactName = GetAttributeValue( options.FileFormat, AttributeKey.OriginContactName );
            var contactPhone = GetAttributeValue( options.FileFormat, AttributeKey.OriginContactPhone );

            int cashHeaderId = GetSystemSetting( SystemSettingNextCashHeaderId ).AsIntegerOrNull() ?? 0;

            var header = new Records.X937.CashLetterHeader
            {
                ID = cashHeaderId.ToString( "D8" ),
                CollectionTypeIndicator = 01,
                DestinationRoutingNumber = destinationRoutingNumber.ToStringSafe(),
                ClientInstitutionRoutingNumber = institutionRoutingNumber.ToStringSafe(),
                BusinessDate = options.BusinessDateTime,
                CreationDateTime = options.ExportDateTime,
                RecordTypeIndicator = "I",
                DocumentationTypeIndicator = "G",
                OriginatorContactName = contactName,
                OriginatorContactPhoneNumber = contactPhone
            };
            SetSystemSetting( SystemSettingNextCashHeaderId, ( cashHeaderId + 1 ).ToString() );

            return header;
        }


        /// <summary>
        /// Gets the cash letter control record (type 90).
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="records">Existing records in the cash letter.</param>
        /// <returns>A CashLetterControl record.</returns>
        protected virtual Records.X937.CashLetterControl GetCashLetterControlRecord( ExportOptions options, List<Record> records )
        {
            var bundleHeaderRecords = records.Where( r => r.RecordType == 20 );
            var checkDetailRecords = records.Where( r => r.RecordType == 25 ).Cast<dynamic>();

            //count deposit records as well as items
            //var itemRecords = records.Where( r => r.RecordType == 25 ); // Only count checks!
            var itemRecords = records.Where( r => r.RecordType == 25 || r.RecordType == 61 );

            var imageDetailRecords = records.Where( r => r.RecordType == 52 );
            var institutionName = GetValueWithFallback( options, AttributeKey.InstitutionName, AttributeKey.OriginName );

            var control = new Records.X937.CashLetterControl
            {
                BundleCount = bundleHeaderRecords.Count(),
                ItemCount = itemRecords.Count(),
                TotalAmount = checkDetailRecords.Sum( c => ( decimal ) c.ItemAmount ),
                ImageCount = imageDetailRecords.Count(),
                ECEInstitutionName = institutionName,
                SettlementDate = options.ExportDateTime
            };

            return control;
        }

        #endregion

        #region Bundle Records

        /// <summary>
        /// Gets all the bundle records in required for the transactions specified.
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="transactions">The transactions to be exported.</param>
        /// <returns>A collection of records that identify all the exported transactions.</returns>
        protected virtual List<Record> GetBundleRecords( ExportOptions options, List<FinancialTransaction> transactions )
        {
            var records = new List<Record>();

            for ( int bundleIndex = 0; ( bundleIndex * MaxItemsPerBundle ) < transactions.Count(); bundleIndex++ )
            {
                var bundleRecords = new List<Record>();
                var bundleTransactions = transactions.Skip( bundleIndex * MaxItemsPerBundle )
                    .Take( MaxItemsPerBundle )
                    .ToList();

                //
                // Add the bundle header for this set of transactions.
                //
                bundleRecords.Add( GetBundleHeader( options, bundleIndex ) );

                //
                // Allow subclasses to provide credit detail records (type 61) if they want.
                //
                bundleRecords.AddRange( GetCreditDetailRecords( options, bundleIndex, bundleTransactions ) );

                //
                // Add records for each transaction in the bundle.
                //
                foreach ( var transaction in bundleTransactions )
                {
                    try
                    {
                        bundleRecords.AddRange( GetItemRecords( options, transaction ) );
                    }
                    catch ( Exception ex )
                    {
                        //Could be bad MICR or Amount; erase so we can re-run.
                        using ( Rock.Data.RockContext rockContext = new Rock.Data.RockContext() )
                        {
                            var badTransation = new FinancialTransactionService( rockContext ).Get( transaction.Id );
                            badTransation.MICRStatus = null;
                            badTransation.CheckMicrEncrypted = string.Empty;
                            badTransation.CheckMicrHash = string.Empty;
                            badTransation.CheckMicrParts = string.Empty;
                            rockContext.SaveChanges();
                        }
                        throw new Exception( string.Format( "Error processing transaction {0}. ({1})", transaction.Id, ex.Message ), ex );
                    }
                }

                //
                // Add the bundle control record.
                //
                bundleRecords.Add( GetBundleControl( options, bundleRecords ) );

                records.AddRange( bundleRecords );
            }

            return records;
        }

        /// <summary>
        /// Gets the bundle header record (type 20).
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="bundleIndex">Number of existing bundle records in the cash letter.</param>
        /// <returns>A BundleHeader record.</returns>
        protected virtual Records.X937.BundleHeader GetBundleHeader( ExportOptions options, int bundleIndex )
        {
            string destinationRoutingNumber = GetValueWithFallback( options, AttributeKey.DestinationRoutingNumber, AttributeKey.ObsoleteRoutingNumber );
            string institutionRoutingNumber = GetValueWithFallback( options, AttributeKey.InstitutionRoutingNumber, AttributeKey.ObsoleteRoutingNumber  );
            if ( institutionRoutingNumber.IsNullOrWhiteSpace() )
            {
                institutionRoutingNumber = destinationRoutingNumber;
            }

            var header = new Records.X937.BundleHeader
            {
                CollectionTypeIndicator = 1,
                DestinationRoutingNumber = destinationRoutingNumber,
                ClientInstitutionRoutingNumber = institutionRoutingNumber,
                BusinessDate = options.BusinessDateTime,
                CreationDate = options.ExportDateTime,
                ID = ( bundleIndex + 1 ).ToString(),
                SequenceNumber = ( bundleIndex + 1 ).ToString(),
                CycleNumber = string.Empty,
                ReturnLocationRoutingNumber = destinationRoutingNumber
            };

            return header;
        }

        /// <summary>
        /// Gets the credit detail deposit record (type 61).
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="bundleIndex">Number of existing bundle records in the cash letter.</param>
        /// <param name="transactions">The transactions associated with this deposit.</param>
        /// <returns>A collection of records.</returns>
        protected virtual List<Record> GetCreditDetailRecords( ExportOptions options, int bundleIndex, List<FinancialTransaction> transactions )
        {
            string originRoutingNumber = GetValueWithFallback( options, AttributeKey.OriginRoutingNumber, AttributeKey.ObsoleteAccountNumber );
            string institutionRoutingNumber = GetValueWithFallback( options, AttributeKey.InstitutionRoutingNumber, AttributeKey.ObsoleteRoutingNumber );
            var creditDetailRecordType = GetAttributeValue( options.FileFormat, AttributeKey.CreditRecordType ).ConvertToEnum<CreditDetailRecordType>( CreditDetailRecordType.None );
            var creditDepositCheckNumber = GetAttributeValue( options.FileFormat, AttributeKey.CreditDepositCheckNumber );

            var records = new List<Record>();

            if ( creditDetailRecordType != CreditDetailRecordType.None )
            {
                var itemAmount = transactions.Sum( p => p.TotalAmount );
                var sequenceNumber = GetNextItemSequenceNumber().ToString( "000000000000000" );

                if ( creditDetailRecordType == CreditDetailRecordType.Type61 )
                {
                    CreditDetail creditDetail = new CreditDetail
                    {
                        AuxiliaryOnUs = string.Empty,
                        ExternalProcessingCode = string.Empty,
                        PayorRoutingNumber = institutionRoutingNumber,
                        CreditAccountNumber = originRoutingNumber + "/" + creditDepositCheckNumber,
                        Amount = itemAmount,
                        InstitutionItemSequenceNumber = sequenceNumber, // A number assigned by you that uniquely identifies the item in the cash letter
                        DocumentTypeIndicator = "G", // Field value must be "G" - Meaning there are 2 images present.
                        DebitCreditIndicator = "2"

                    };
                    records.Add( creditDetail );
                }

                if ( creditDetailRecordType == CreditDetailRecordType.Type61A )
                {
                    CreditReconciliation creditReconciliation = new CreditReconciliation
                    {
                        RecordUsageIndicator = 5,
                        AuxiliaryOnUs = string.Empty,
                        ExternalProcessingCode = string.Empty,
                        PostingAccountRoutingNumber = institutionRoutingNumber.AsInteger(),
                        PostingAccountBankOnUs = originRoutingNumber + "/" + creditDepositCheckNumber,
                        ItemAmount = itemAmount,
                        ECEInstitutionSequenceNumber = sequenceNumber, // A number assigned by you that uniquely identifies the item in the cash letter
                        DocumentationTypeIndicator = "G", // Field value must be "G" - Meaning there are 2 images present.
                                                          //TypeOfAccountCode = string.Empty,
                                                          //SourceOfWork = string.Empty, // Field value must be "2" space or "02" meaning internal - branch
                    };
                    records.Add( creditReconciliation );
                }

                for ( int i = 0; i < 2; i++ )
                {
                    using ( var ms = GetDepositSlipImage( options, itemAmount, i == 0, transactions ) )
                    {
                        //
                        // Get the Image View Detail record (type 50).
                        //
                        var detail = new ImageViewDetail
                        {
                            ImageIndicator = 1,
                            ImageCreatorRoutingNumber = institutionRoutingNumber,
                            ImageCreatorDate = options.ExportDateTime,
                            ImageViewFormatIndicator = 0,
                            CompressionAlgorithmIdentifier = 0,
                            SideIndicator = i,
                            ViewDescriptor = 0,
                            DigitalSignatureIndicator = 0,
                            DataSize = ( int ) ms.Length
                        };

                        //
                        // Get the Image View Data record (type 52).
                        //
                        var data = new ImageViewData
                        {
                            InstitutionRoutingNumber = institutionRoutingNumber,
                            BundleBusinessDate = options.BusinessDateTime,
                            ClientInstitutionItemSequenceNumber = sequenceNumber,
                            ClippingOrigin = 0,
                            ImageData = ms.ReadBytesToEnd()
                        };

                        records.Add( detail );
                        records.Add( data );
                    }
                }
            }


            return records;
        }

        /// <summary>
        /// Gets the credit detail deposit record (type 61).
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="transactions">The transactions associated with this deposit.</param>
        /// <param name="isFrontSide">True if the image to be retrieved is the front image.</param>
        /// <returns>A stream that contains the image data in TIFF 6.0 CCITT Group 4 format.</returns>
        protected virtual Stream GetDepositSlipImage( ExportOptions options, Decimal itemAmount, bool isFrontSide, List<FinancialTransaction> transactions )
        {
            var bitmap = new System.Drawing.Bitmap( 1200, 550 );
            var g = System.Drawing.Graphics.FromImage( bitmap );

            var depositSlipTemplate = GetAttributeValue( options.FileFormat, AttributeKey.DepositSlipTemplate );
            var mergeFields = new Dictionary<string, object>
            {
                { "FileFormat", options.FileFormat },
                { "Amount", itemAmount.ToString( "C" ) },
                { "ItemCount", transactions.Count() }
            };
            var depositSlipText = depositSlipTemplate.ResolveMergeFields( mergeFields, null );

            //
            // Ensure we are opague with white.
            //
            g.FillRectangle( System.Drawing.Brushes.White, new System.Drawing.Rectangle( 0, 0, 1200, 550 ) );

            if ( isFrontSide )
            {
                g.DrawString( depositSlipText,
                    new System.Drawing.Font( "Tahoma", 30 ),
                    System.Drawing.Brushes.Black,
                    new System.Drawing.PointF( 50, 50 ) );
            }

            g.Flush();

            //
            // Ensure the DPI is correct.
            //
            bitmap.SetResolution( 200, 200 );

            //
            // Compress using TIFF, CCITT Group 4 format.
            //
            var codecInfo = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                .Where( c => c.MimeType == "image/tiff" )
                .First();
            var parameters = new System.Drawing.Imaging.EncoderParameters( 1 );
            parameters.Param[0] = new System.Drawing.Imaging.EncoderParameter( System.Drawing.Imaging.Encoder.Compression, ( long ) System.Drawing.Imaging.EncoderValue.CompressionCCITT4 );

            var ms = new MemoryStream();
            bitmap.Save( ms, codecInfo, parameters );
            ms.Position = 0;

            return ms;
        }

        /// <summary>
        /// Gets the bundle control record (type 70).
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="records">The existing records in the bundle.</param>
        /// <returns>A BundleControl record.</returns>
        protected virtual Records.X937.BundleControl GetBundleControl( ExportOptions options, List<Record> records )
        {
            var itemRecords = records.Where( r => r.RecordType == 25 || r.RecordType == 61 );
            var checkDetailRecords = records.Where( r => r.RecordType == 25 ).Cast<dynamic>();
            var imageDetailRecords = records.Where( r => r.RecordType == 52 );

            var control = new Records.X937.BundleControl
            {
                ItemCount = itemRecords.Count(),
                TotalAmount = checkDetailRecords.Sum( r => ( decimal ) r.ItemAmount ),
                MICRValidTotalAmount = checkDetailRecords.Sum( r => ( decimal ) r.ItemAmount ),
                ImageCount = imageDetailRecords.Count()
            };

            return control;
        }

        #endregion

        #region Item Records

        /// <summary>
        /// Gets the records that identify a single check being deposited.
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="transaction">The transaction to be deposited.</param>
        /// <returns>A collection of records.</returns>
        protected virtual List<Record> GetItemRecords( ExportOptions options, FinancialTransaction transaction )
        {
            var records = new List<Record>();

            records.AddRange( GetItemDetailRecords( options, transaction ) );

            try
            {
                records.AddRange( GetImageRecords( options, transaction, transaction.Images.Take( 1 ).First(), true ) );
                records.AddRange( GetImageRecords( options, transaction, transaction.Images.Skip( 1 ).Take( 1 ).First(), false ) );
            }
            catch ( Exception ex )
            {
                throw new ArgumentException( string.Format( "Transaction Does Not Contain Two Valid Image Records (Front and Back Check Scans)", transaction.Id ), ex );
            }

            return records;
        }

        /// <summary>
        /// Gets the item detail records (type 25, 26, etc.)
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="transaction">The transaction being deposited.</param>
        /// <returns>A collection of records.</returns>
        protected virtual List<Record> GetItemDetailRecords( ExportOptions options, FinancialTransaction transaction )
        {
            string originRoutingNumber = GetValueWithFallback( options, AttributeKey.OriginRoutingNumber, AttributeKey.ObsoleteAccountNumber );
            string institutionRoutingNumber = GetValueWithFallback( options, AttributeKey.InstitutionRoutingNumber, AttributeKey.ObsoleteRoutingNumber );
            string bofdRoutingNumber = Rock.Security.Encryption.DecryptString( GetAttributeValue( options.FileFormat, AttributeKey.BOFDRoutingNumber ) );
            var sequenceNumber = GetNextItemSequenceNumber();


            if ( bofdRoutingNumber.IsNullOrWhiteSpace() )
            {
                if ( institutionRoutingNumber.IsNullOrWhiteSpace() )
                {
                    institutionRoutingNumber = originRoutingNumber;
                }

                bofdRoutingNumber = institutionRoutingNumber;
            }

            var institutionSequenceNumber = sequenceNumber.ToString();
            var sequenceNumberJustification = GetAttributeValue( options.FileFormat, AttributeKey.ItemSequenceNumberJustification );
            if ( sequenceNumberJustification == "Left" )
            {
                institutionSequenceNumber = institutionSequenceNumber.PadRight( 15, ' ' ).Left( 15 );
            }

            //
            // Parse the MICR data from the transaction.
            //
            var micr = GetMicrInstance( transaction.CheckMicrEncrypted );

            var transactionRoutingNumber = micr.GetRoutingNumber();

            //
            // On-Us Calculation (check account numbers can be too big)
            //
            string onUs = string.Format( "{0}/{1}", micr.GetAccountNumber().TrimStart( '0' ), micr.GetCheckNumber().TrimStart( '0' ) );
            if ( onUs.Length > 20 ) //too big
            {
                int checkNumberLength = onUs.Length - micr.GetAccountNumber().Length - 1;  //number of characters left for the check number
                if ( checkNumberLength < 3 )
                {
                    checkNumberLength = 3; //Minimum length of 3
                }
                onUs = string.Format( "{0}/{1}", micr.GetAccountNumber().TrimStart( '0' ), micr.GetCheckNumber().TrimStart( '0' ).Right( checkNumberLength ) );
            }

            //
            // Get the Check Detail record (type 25).
            //
            var detail = new Records.X937.CheckDetail
            {
                PayorBankRoutingNumber = transactionRoutingNumber.Substring( 0, 8 ),
                PayorBankRoutingNumberCheckDigit = transactionRoutingNumber.Substring( 8, 1 ),
                OnUs = onUs.Right( 20 ), //get just right side of field as the front of an account number may be chopped off
                ExternalProcessingCode = micr.GetExternalProcessingCode(),
                AuxiliaryOnUs = micr.GetAuxOnUs(),
                ItemAmount = transaction.TotalAmount,
                ClientInstitutionItemSequenceNumber = institutionSequenceNumber,
                DocumentationTypeIndicator = "G",
                MICRValidIndicator = 1,
                BankOfFirstDepositIndicator = "Y",
                CheckDetailRecordAddendumCount = 1
            };

            //
            // Get the Addendum A record (type 26).
            //
            string truncationIndicator = GetAttributeValue( options.FileFormat, AttributeKey.TruncationIndicator ).Substring( 0, 1 ).ToUpper();
            var detailA = new Records.X937.CheckDetailAddendumA
            {
                RecordNumber = 1,
                BankOfFirstDepositRoutingNumber = bofdRoutingNumber,
                BankOfFirstDepositBusinessDate = options.BusinessDateTime,
                TruncationIndicator = truncationIndicator,
                BankOfFirstDepositItemSequenceNumber = sequenceNumber.ToString( "000000000000000" ),
                BankOfFirstDepositConversionIndicator = "2",
                BankOfFirstDepositCorrectionIndicator = "0"
            };

            return new List<Record> { detail, detailA };
        }

        /// <summary>
        /// Gets the image record for a specific transaction image (type 50 and 52).
        /// </summary>
        /// <param name="options">Export options to be used by the component.</param>
        /// <param name="transaction">The transaction being deposited.</param>
        /// <param name="image">The check image scanned by the scanning application.</param>
        /// <param name="isFront">if set to <c>true</c> [is front].</param>
        /// <returns>A collection of records.</returns>
        protected virtual List<Record> GetImageRecords( ExportOptions options, FinancialTransaction transaction, FinancialTransactionImage image, bool isFront )
        {
            string destinationRoutingNumber = GetValueWithFallback( options, AttributeKey.DestinationRoutingNumber, AttributeKey.ObsoleteRoutingNumber );
            string originRoutingNumber = GetValueWithFallback( options, AttributeKey.OriginRoutingNumber, AttributeKey.ObsoleteAccountNumber );
            string institutionRoutingNumber = GetValueWithFallback( options, AttributeKey.InstitutionRoutingNumber, AttributeKey.ObsoleteRoutingNumber );
            if ( institutionRoutingNumber.IsNullOrWhiteSpace() )
            {
                institutionRoutingNumber = originRoutingNumber;
            }

            var institutionSequenceNumber = GetNextItemSequenceNumber().ToString();
            var sequenceNumberJustification = GetAttributeValue( options.FileFormat, AttributeKey.ItemSequenceNumberJustification );
            if ( sequenceNumberJustification == "Left" )
            {
                institutionSequenceNumber = institutionSequenceNumber.PadRight( 15, ' ' ).Left( 15 );
            }

            var checkEndorsement = GetAttributeValue( options.FileFormat, AttributeKey.CheckEndorsementTemplate );
            var enableEndorsement = GetAttributeValue( options.FileFormat, AttributeKey.EnableDigitalEndorsement ).AsBoolean();


            //
            // If endorsement, add to back of image
            //
            Stream imageData = image.BinaryFile.ContentStream;

            if ( !isFront && enableEndorsement && checkEndorsement.IsNotNullOrWhiteSpace() ) //if image is the back of the check, add endorsement
            {
                //add endorsement for back image
                var bitmap = new Bitmap( image.BinaryFile.ContentStream );
                bitmap.SetResolution( 200, 200 );
                var newBitmap = new Bitmap( bitmap.Width, bitmap.Height );
                newBitmap.SetResolution( 200, 200 );
                var g = System.Drawing.Graphics.FromImage( newBitmap );
                g.DrawImage( bitmap, 0, 0 ); //draw image onto blank bitmap

                var mergeFields = new Dictionary<string, object>
                {
                    { "FileFormat", options.FileFormat },
                    { "Amount", transaction.TotalAmount.ToString( "C" ) },
                    { "BusinessDate", options.BusinessDateTime }
                };
                var checkEndorsementText = checkEndorsement.ResolveMergeFields( mergeFields, null );

                g.DrawString( checkEndorsementText,
                    new System.Drawing.Font( "Tahoma", 10 ),
                    System.Drawing.Brushes.Black,
                    new System.Drawing.PointF( ( bitmap.Width / 2 ) - 40, ( bitmap.Height / 2 ) - 40 ) );

                g.Flush();

                //
                // Ensure the DPI is correct.
                //

                imageData = new System.IO.MemoryStream();
                newBitmap.Save( imageData, System.Drawing.Imaging.ImageFormat.Tiff );//maybe this works...

            }

            var tiffImageBytes = ConvertImageToTiffG4( imageData ).ReadBytesToEnd();
            //
            // Get the Image View Detail record (type 50).
            //
            var detail = new Records.X937.ImageViewDetail
            {
                ImageIndicator = 1,
                ImageCreatorRoutingNumber = destinationRoutingNumber,
                ImageCreatorDate = image.CreatedDateTime ?? options.ExportDateTime,
                ImageViewFormatIndicator = 0,
                CompressionAlgorithmIdentifier = 0,
                SideIndicator = isFront ? 0 : 1,
                ViewDescriptor = 0,
                DigitalSignatureIndicator = 0,
                DataSize = ( int ) tiffImageBytes.Length,
                DigitalSignatureMethod = 0,
                SecurityKeySize = 00000,
                StartOfProtectedData = 0000000,
                LengthOfProtectedData = 0000000
            };

            //
            // Get the Image View Data record (type 52).
            //
            var data = new Records.X937.ImageViewData
            {
                InstitutionRoutingNumber = institutionRoutingNumber,
                BundleBusinessDate = options.BusinessDateTime,
                ClientInstitutionItemSequenceNumber = institutionSequenceNumber,
                ClippingOrigin = 0,
                ClippingCoordinateH1 = 0000,
                ClippingCoordinateH2 = 0000,
                ClippingCoordinateV1 = 0000,
                ClippingCoordinateV2 = 0000,
                ImageData = tiffImageBytes
            };

            return new List<Record> { detail, data };
        }


        #endregion

        #region Protected Methods

        /// <summary>
        /// Converts the image to tiff g4 specifications.
        /// </summary>
        /// <param name="imageStream">The image stream.</param>
        /// <returns></returns>
        protected Stream ConvertImageToTiffG4( Stream imageStream )
        {
            var bitmap = new Bitmap( imageStream );

            //
            // Ensure the DPI is correct.
            //
            bitmap.SetResolution( 200, 200 );

            //
            // Compress using TIFF, CCITT Group 4 format.
            //
            var codecInfo = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                .Where( c => c.MimeType == "image/tiff" )
                .First();
            var parameters = new System.Drawing.Imaging.EncoderParameters( 1 );
            parameters.Param[0] = new System.Drawing.Imaging.EncoderParameter( System.Drawing.Imaging.Encoder.Compression, ( long ) System.Drawing.Imaging.EncoderValue.CompressionCCITT4 );

            var ms = new MemoryStream();
            bitmap.Save( ms, codecInfo, parameters );
            ms.Position = 0;

            return ms;
        }

        /// <summary>
        /// Gets a MICR object instance from the encrypted MICR data.
        /// </summary>
        /// <param name="encryptedMicrContent">Content of the encrypted MICR.</param>
        /// <returns>A <see cref="Micr"/> instance that can be used to get decrypted MICR data.</returns>
        /// <exception cref="ArgumentException">MICR data is empty.</exception>
        protected Micr GetMicrInstance( string encryptedMicrContent )
        {
            string decryptedMicrContent = Rock.Security.Encryption.DecryptString( encryptedMicrContent );

            if ( decryptedMicrContent == null )
            {
                throw new ArgumentException( "MICR data is empty." );
            }

            var micr = new Micr( decryptedMicrContent );

            return micr;
        }

        protected string GetValueWithFallback( ExportOptions options, string currentKey, string oldKey )
        {
            var attributeValue = Rock.Security.Encryption.DecryptString( GetAttributeValue( options.FileFormat, currentKey ) );
            if ( attributeValue.IsNullOrWhiteSpace() )
            {
                attributeValue = Rock.Security.Encryption.DecryptString( GetAttributeValue( options.FileFormat, oldKey ) );
            }

            return attributeValue;
        }

        #endregion
    }

    public enum CreditDetailRecordType
    {
        None = 0,
        Type61 = 1,
        Type61A = 2
    }
}
