using Autodesk.AutoCAD.DatabaseServices;
namespace Blocks_for_Nesting
{
    public static class TransactionExtensions
    {
        // A simple extension method that aggregates the extents of any entities
        // from Through the interface author Kean Walmsley
        public static Extents3d GetExtents(this Transaction tr, ObjectIdCollection ids)
        {
            var ext = new Extents3d();
            foreach (ObjectId id in ids)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent != null)
                { ext.AddExtents(ent.GeometricExtents); }
            }
            return ext;
        }
    }
}
