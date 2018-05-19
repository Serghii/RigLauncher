
namespace RigLauncher
{
    public class RigVersion
    {
        public RigVersion(int curVersion, int newVersion, string zipId)
        {
            this.curVersion = curVersion;
            this.newVersion = newVersion;
            ZipId = zipId;
        }

        public int  curVersion { get; }
        public int  newVersion { get; }
        public string  ZipId { get; }
    }
}
