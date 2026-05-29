using System;
using System.Diagnostics;

namespace techlink_workspace.Controller.Logic.IDGenerate
{
    /// <summary>
    /// Pure C# ID generator — no IKVM / Java dependency.
    /// Produces the same style of ascending alphanumeric IDs as the original.
    /// </summary>
    public class UUIDGenerator
    {
        private static readonly long START_TMP = 1465142400000L;
        private static readonly string PID;            // process id, resolved once at startup
        private static long sequence = 0L;
        private static long userSequence = 0L;
        private static long tenantSequence = 0L;
        private static long last_tmp = -1L;
        private static long lastUserId = -1L;
        private static long lastTenantId = -1L;
        private static readonly object _lock = new object();

        private UUIDGenerator() { }

        // ── Use Process.GetCurrentProcess() instead of IKVM ManagementFactory ──
        static UUIDGenerator()
        {
            PID = Process.GetCurrentProcess().Id.ToString();
        }

        // ── Time helpers ─────────────────────────────────────────────────────
        private static long getCurrentTime()
            => DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();

        private static long getDiffTime()
        {
            long diff = getCurrentTime() - START_TMP;
            TimeSpan ts = TimeSpan.FromMilliseconds(diff);
            DateTime dt = new DateTime(1970, 1, 1) + ts;
            string s = dt.ToString("HHmmssfff");
            return long.Parse(diff / 86400000L + s);
        }

        // ── Public API ───────────────────────────────────────────────────────
        public static string getId()
        {
            lock (_lock)
            {
                long sec = getDiffTime();
                sequence = (sec == last_tmp) ? sequence + 1 : 0L;
                last_tmp = sec;
                return sec * 10000L + sequence + PID;
            }
        }

        public static string getAscId()
        {
            lock (_lock)
            {
                long sec = getDiffTime();
                sequence = (sec == last_tmp) ? sequence + 1 : 0L;
                last_tmp = sec;
                return HexTransformatUtil.hex10ToAnly(sec * 10000 + sequence)
                     + HexTransformatUtil.hex10ToAnly(long.Parse(PID));
            }
        }

        public static string getUserId()
        {
            lock (_lock)
            {
                long sec = getDiffTime() / 1000L;
                userSequence = (sec == lastUserId) ? userSequence + 1 : 0L;
                lastUserId = sec;
                return HexTransformatUtil.hex10ToAnly(sec)
                     + HexTransformatUtil.hex10ToAnly(userSequence)
                     + HexTransformatUtil.hex10ToAnly(long.Parse(PID));
            }
        }

        public static string getTenantId()
        {
            lock (_lock)
            {
                long sec = getDiffTime() / 100000L;
                tenantSequence = (sec == lastTenantId) ? tenantSequence + 1 : 0L;
                lastTenantId = sec;
                return HexTransformatUtil.hex10ToAnly(sec)
                     + HexTransformatUtil.hex10ToAnly(tenantSequence)
                     + HexTransformatUtil.hex10ToAnly(long.Parse(PID));
            }
        }
    }
}