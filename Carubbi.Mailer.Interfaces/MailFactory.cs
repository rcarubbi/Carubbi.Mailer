using Carubbi.Mailer.Interfaces;
using Carubbi.Utils.IoC;

namespace Carubbi.Mailer.Factories
{
    public class MailFactory
    {
        private static volatile MailFactory _instance;
        private static volatile object _locker = new object();

        private MailFactory()
        { }


        public static MailFactory GetInstance()
        {
            if (_instance == null)
            {
                lock (_locker)
                {
                    if (_instance == null)
                    {
                        _instance = new MailFactory();
                    }
                }
            }
            return _instance;

        }

        public IMailSender CreateSender()
        {
            return ImplementationResolver.Resolve<IMailSender>();
        }

        public IMailReceiver CreateReceiver()
        {
            return ImplementationResolver.Resolve<IMailReceiver>();
        }

    }
}
