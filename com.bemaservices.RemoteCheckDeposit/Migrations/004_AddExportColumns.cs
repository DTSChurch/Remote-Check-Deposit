using Rock.Plugin;

namespace com.bemaservices.RemoteCheckDeposit.Migrations
{
    [MigrationNumber( 4, "1.12.7" )]
    public class AddExportColumns : Migration
    {
        public override void Up()
        {
            // Entity: Rock.Model.FinancialBatch Attribute: Export Date
            RockMigrationHelper.AddOrUpdateEntityAttribute( "Rock.Model.FinancialBatch", "FE95430C-322D-4B67-9C77-DFD1D4408725", "", "", "Export Date", "Export Date", @"", 0, @"", SystemGuid.Attribute.FINANCIAL_BATCH_EXPORT_DATE, "com.bemaservices.ExportDate" );
            // Entity: Rock.Model.FinancialBatch Attribute: Export File
            RockMigrationHelper.AddOrUpdateEntityAttribute( "Rock.Model.FinancialBatch", "6F9E2DD0-E39E-4602-ADF9-EB710A75304A", "", "", "Export File", "Export File", @"", 0, @"", SystemGuid.Attribute.FINANCIAL_BATCH_EXPORT_FILE, "com.bemaservices.ExportFile" );

            // Qualifier for attribute: com.bemaservices.ExportDate
            RockMigrationHelper.UpdateAttributeQualifier( SystemGuid.Attribute.FINANCIAL_BATCH_EXPORT_DATE, "format", @"", "AE10767A-CB0E-4174-A08C-65E1D8617271" );
            // Qualifier for attribute: com.bemaservices.ExportDate
            RockMigrationHelper.UpdateAttributeQualifier( SystemGuid.Attribute.FINANCIAL_BATCH_EXPORT_DATE, "displayDiff", @"False", "C161653F-27E1-4F88-90D1-0DFB5B078AD2" );
            // Qualifier for attribute: com.bemaservices.ExportDate
            RockMigrationHelper.UpdateAttributeQualifier( SystemGuid.Attribute.FINANCIAL_BATCH_EXPORT_DATE, "displayCurrentOption", @"False", "24747662-74F1-423C-84C0-8B8AB5EFC223" );
            // Qualifier for attribute: com.bemaservices.ExportDate
            RockMigrationHelper.UpdateAttributeQualifier( SystemGuid.Attribute.FINANCIAL_BATCH_EXPORT_DATE, "datePickerControlType", @"Date Picker", "EB572C27-5184-4CB1-967D-E28821047BFE" );
            // Qualifier for attribute: com.bemaservices.ExportDate
            RockMigrationHelper.UpdateAttributeQualifier( SystemGuid.Attribute.FINANCIAL_BATCH_EXPORT_DATE, "futureYearCount", @"", "F362F913-479C-4752-A7D3-ECAF746618DB" );
            // Qualifier for attribute: com.bemaservices.ExportFile
            RockMigrationHelper.UpdateAttributeQualifier( SystemGuid.Attribute.FINANCIAL_BATCH_EXPORT_FILE, "binaryFileType", Rock.SystemGuid.BinaryFiletype.DEFAULT, "16119D35-D571-480F-84E1-2DA2F078A429" );

        }

        public override void Down()
        {
            RockMigrationHelper.DeleteAttribute( SystemGuid.Attribute.FINANCIAL_BATCH_EXPORT_DATE ); // Rock.Model.FinancialBatch: Export Date
            RockMigrationHelper.DeleteAttribute( SystemGuid.Attribute.FINANCIAL_BATCH_EXPORT_FILE ); // Rock.Model.FinancialBatch: Export File

        }
    }
}
