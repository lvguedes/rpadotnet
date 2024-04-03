using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using RpaLib.ProcessAutomation;
using System.Security.Authentication;

namespace RpaLibXunit
{
    public class CredentialManager
    {
        [Fact]
        public void GetSavedCredential()
        {
            var target = "mycredtest";
            var mycred = Ut.GetStoredCredential(target);

            //Console.WriteLine("Testing credential target: " + target);
            //Console.WriteLine("The username is: " + mycred.Username);
            //Console.WriteLine("The password is: " + mycred.Password);

            Assert.Equal("justauser", mycred.Username);
            Assert.Equal("123456", mycred.Password);
        }

        [Fact]
        public void GetWrongCredential()
        {
            var target = "badCredential";

            try
            {
                var mycred = Ut.GetStoredCredential(target);
            }
            catch (Exception ex)
            {
                Assert.IsType<InvalidCredentialException>(ex);
            }
        }
    }
}
