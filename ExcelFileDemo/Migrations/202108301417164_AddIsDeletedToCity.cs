namespace ExcelFileDemo.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddIsDeletedToCity : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.Cities", "IsDeleted", c => c.Boolean(nullable: false));
        }
        
        public override void Down()
        {
            DropColumn("dbo.Cities", "IsDeleted");
        }
    }
}
