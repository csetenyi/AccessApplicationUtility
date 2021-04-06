

using System;
using ac = AccessApplication;

namespace AccessApplicationCaller
{
    class Program
    {

        static void Main(string[] args)
        {

            var ace = new ac.AccessDatabase("C:\\Users\\Csetényi Zoltán\\Documents\\Test\\test.accdb");

            ace.RunSQLStatementDoCmd("SELECT * INTO TableNew FROM Table1;");



        }
    }
}
