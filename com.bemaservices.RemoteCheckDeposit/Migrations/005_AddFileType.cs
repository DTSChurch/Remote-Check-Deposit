using Rock.Plugin;

namespace com.bemaservices.RemoteCheckDeposit.Migrations
{
    [MigrationNumber( 5, "1.12.7" )]
    public class AddFileType : Migration
    {
        public override void Up()
        {
            RockMigrationHelper.UpdateBinaryFileTypeRecord( "0AA42802-04FD-4AEC-B011-FEB127FC85CD", "Remote Check Deposit File", "Generated Remote Deposit files for Banks", "fa fa-money", SystemGuid.BinaryFileType.REMOTE_CHECK_DEPOSIT, false, true );
            RockMigrationHelper.AddSecurityAuthForBinaryFileType( SystemGuid.BinaryFileType.REMOTE_CHECK_DEPOSIT,3, "View", false, null, Rock.Model.SpecialRole.AllUsers, "0EF64485-976E-4F54-929A-7CDFBB542F84" );
            RockMigrationHelper.AddSecurityAuthForBinaryFileType( SystemGuid.BinaryFileType.REMOTE_CHECK_DEPOSIT,0, "View", true, Rock.SystemGuid.Group.GROUP_ADMINISTRATORS, Rock.Model.SpecialRole.None, "05B582E9-508F-4538-9EF0-BEB1940D895F" );
            RockMigrationHelper.AddSecurityAuthForBinaryFileType( SystemGuid.BinaryFileType.REMOTE_CHECK_DEPOSIT,1,"View", true, Rock.SystemGuid.Group.GROUP_FINANCE_ADMINISTRATORS, Rock.Model.SpecialRole.None, "7173D8A7-5983-4A9D-8BD5-6219577697DC" );
            RockMigrationHelper.AddSecurityAuthForBinaryFileType( SystemGuid.BinaryFileType.REMOTE_CHECK_DEPOSIT,2,"View", true, Rock.SystemGuid.Group.GROUP_FINANCE_USERS, Rock.Model.SpecialRole.None, "7D19CC09-CDB6-4A2F-84A3-91C4DB530C0A" );
        }

        public override void Down()
        {
        }
    }
}
