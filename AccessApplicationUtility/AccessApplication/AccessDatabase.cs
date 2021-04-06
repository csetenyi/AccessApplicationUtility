using System;
// using System.Runtime.InteropServices;
using Access = Microsoft.Office.Interop.Access;
using Dao = Microsoft.Office.Interop.Access.Dao;
// using Microsoft.Office.Interop.Access.Dao;
// using System.Data;
// using System.Xml;
// using System.Drawing;
// using System.Data;

namespace AccessApplication
{

    public enum TypeOfAccessDatabaseCollection
    {
        QueryDef = 1,
        ImportSpecification = 2
    }
    public class AccessDatabase
    {

        // PROPERTIES - application

        Access.Application oAccess = null;
        Dao.Database oDatabase = null;

        // LIFECYCLE Methods
        public AccessDatabase(String FileFullPath)
        {

            try
            {
                oAccess = new Access.Application();
                oAccess.Visible = true;
                oAccess.OpenCurrentDatabase(FileFullPath, true);
                oDatabase = oAccess.CurrentDb();
                Console.WriteLine("Access database is opened.");
            }
            catch
            {
                Console.WriteLine("Failed to open file");
            }
        }

        public AccessDatabase(String FileFullPath, String Password)
        {

            try
            {
                oAccess = new Access.Application();
                oAccess.Visible = true;
                oAccess.OpenCurrentDatabase(FileFullPath, true, Password);
                oDatabase = oAccess.CurrentDb();
                Console.WriteLine("Access database is opened.");
            }
            catch
            {
                Console.WriteLine("Failed to open file");
            }
        }

        public String CloseDatabase()
        {
            try
            {
                oAccess.DoCmd.CloseDatabase();
                return "success";
            }
            catch
            {
                return "failed";
            }
            finally
            {
            }
        }

        // UTILITY Methods

        public Boolean ObjectExists(String NameOfTheObject, TypeOfAccessDatabaseCollection ObjectType)
        {

            Object collectionOfObjects;
            Object itemInCollection;
            Boolean objectWithNameFound = false;

            try
            {
                switch (ObjectType)
                {
                    case TypeOfAccessDatabaseCollection.QueryDef:

                        foreach (Dao.QueryDef queryDefinition in oAccess.CurrentDb().QueryDefs)
                        {
                            if (queryDefinition.Name == NameOfTheObject)
                            {
                                objectWithNameFound = true;
                            }
                        }
                        break;

                    case TypeOfAccessDatabaseCollection.ImportSpecification:

                        foreach (Access.ImportExportSpecification importSpecification in oAccess.CurrentProject.ImportExportSpecifications)
                        {
                            if (importSpecification.Name == NameOfTheObject)
                            {
                                objectWithNameFound = true;
                            }
                        }
                        break;

                    default:
                        collectionOfObjects = null;
                        break;
                }





                return objectWithNameFound;
            }
            catch
            {
                return false;
            }
            finally
            {
            }
        }

        // ACTION Methods

        public String RunSQLStatementDoCmd(String sqlStatement)
        {
            try
            {
                oAccess.DoCmd.SetWarnings(false);
                oAccess.DoCmd.RunSQL(sqlStatement);
                oAccess.DoCmd.SetWarnings(true);

                return "success";
            }
            catch
            {
                return "failed";
            }
            finally
            {
            }
        }

        public String GetIndexedQueryName(String BaseQueryName)
        {
            var indexNumber = 0;
            var queryFullIndexedName = "";
            var queryWithIndexAlreadyExist = false;

            do
            {
                indexNumber += 1;
                queryFullIndexedName = $"{BaseQueryName} - {indexNumber.ToString()}";
                queryWithIndexAlreadyExist = ObjectExists(queryFullIndexedName, TypeOfAccessDatabaseCollection.QueryDef);
            } while (queryWithIndexAlreadyExist);


            return queryFullIndexedName;
        }

        public Boolean QueryNameAlreadyExists(String QueryName)
        {
            return ObjectExists(QueryName, TypeOfAccessDatabaseCollection.QueryDef);
        }

        public String CreateNewQueryDefinition(String QueryDefinitionName, String SqlStatement, Boolean AddIndexIfAlreadyExists = true, Boolean FailIfExists = false)
        {
            var newQueryDefinitionName = "";
            var queryDefinitionAlreadyExistsWithThatName = false;

            try
            {
                queryDefinitionAlreadyExistsWithThatName = ObjectExists(QueryDefinitionName, TypeOfAccessDatabaseCollection.QueryDef);

                // Fail if exist:
                if (queryDefinitionAlreadyExistsWithThatName && FailIfExists) { throw new Exception(); }

                // Name has not been used yet.
                if (queryDefinitionAlreadyExistsWithThatName == false)
                {
                    newQueryDefinitionName = QueryDefinitionName;
                }
                // Exist -> Use Index after the name
                else if (queryDefinitionAlreadyExistsWithThatName && AddIndexIfAlreadyExists)
                {
                    newQueryDefinitionName = GetIndexedQueryName(QueryDefinitionName);
                }
                // Delete -> Delete and create new
                else if (queryDefinitionAlreadyExistsWithThatName && AddIndexIfAlreadyExists)
                {
                    oAccess.DoCmd.SetWarnings(false);
                    oAccess.CurrentDb().QueryDefs.Delete(QueryDefinitionName);
                    oAccess.RefreshDatabaseWindow();
                    oAccess.DoCmd.SetWarnings(true);
                    newQueryDefinitionName = QueryDefinitionName; 
                }


                oAccess.DoCmd.SetWarnings(false);
                oAccess.CurrentDb().CreateQueryDef(QueryDefinitionName, SqlStatement);
                oAccess.RefreshDatabaseWindow();
                oAccess.DoCmd.SetWarnings(true);
                return newQueryDefinitionName;
            }
            catch
            {
                return "failure";
            }
            finally
            {
            }
        }

        public String OpenQueryDefinition(String NameOfQueryDefinition)
        {
            try
            {
                oAccess.DoCmd.SetWarnings(false);
                oAccess.DoCmd.OpenQuery(NameOfQueryDefinition);
                oAccess.RefreshDatabaseWindow();
                oAccess.DoCmd.SetWarnings(true);
                return "success";
            }
            catch
            {
                return "fail";
            }
            finally
            {
            }
        }


        public String CurrentDbExecute(String SqlStatement)
        {
            try
            {
                oAccess.CurrentDb().Execute(SqlStatement);
                return "success";
            }
            catch
            {
                return "fail";
            }
        }

        public String RunMacro(String MacroName)
        {
            try
            {
                oAccess.Run(MacroName);
                return "success";
            }
            catch
            {
                return "fail";
            }
        }

        public String ExportData(String TableName, String OutputFileFullPathName, Boolean HasFieldNames, String RangeName)
        {
            try
            {
                oAccess.DoCmd.TransferSpreadsheet(Access.AcDataTransferType.acExport, Access.AcSpreadSheetType.acSpreadsheetTypeExcel12Xml, TableName, OutputFileFullPathName, HasFieldNames, RangeName);
                return "success";
            }
            catch
            {
                return "fail";
            }
        }


        public String RunSavedImportExportSpecification(String SpecificationName, String XmlSpecification)
        {
            try
            {
                oAccess.CurrentProject.ImportExportSpecifications.Add(SpecificationName, XmlSpecification);
                oAccess.DoCmd.RunSavedImportExport(SpecificationName);
                return "success";
            }
            catch
            {
                return "fail";
            }
        }





    }
}
