using Microsoft.VisualStudio.TestPlatform.CommunicationUtilities.Interfaces;
using OneNoteTools;
using System.Runtime.InteropServices;

namespace TestProject
{
    [TestClass]
    public class ToolsLibraryTests
    {

        [TestMethod]
        public void TestConnection()
        {

            Connection conn = null;

            try
            {

                conn = GetConnection();

                // test if connection was successful.
                if (!conn.AppVisible())
                    Assert.Fail("Could not connect to OneNote. Ensure that the OneNote is installed, is running, and a notebook is open.");

                // test current notebook fetch
                Notebook nb = conn.GetCurrentNotebook();
                if(string.IsNullOrEmpty(nb.Name))
                    Assert.Fail("Current notebook retrieval failed. Notebook name is null or empty.");

            }
            catch (Exception ex)
            { Assert.Fail(ex.InnerException.ToString()); }
            finally
            { DisposeConnection(conn); }

        }

        private Connection GetConnection()
        {
            return new Connection();
        }

        private void DisposeConnection(Connection conn)
        {
            if (conn != null)
                conn.Dispose();
        }

    }
}