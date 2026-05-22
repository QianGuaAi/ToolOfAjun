using System;

namespace MyTools.Services
{
    internal static class CosturaBootstrap
    {
        private static readonly object SyncRoot = new object();
        private static bool _initialized;

        public static void EnsureInitialized()
        {
            if (_initialized)
            {
                return;
            }

            lock (SyncRoot)
            {
                if (_initialized)
                {
                    return;
                }

                CosturaUtility.Initialize();
                _initialized = true;
            }
        }
    }
}
