using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateNewRooms
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class SetRoomsTag : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            List<Level> levels = GetLevels(doc);

            RoomTagType roomTag = GetRoomTag(doc);

            using (Transaction transaction = new Transaction(doc, "Create Room"))
            {
                transaction.Start();

                Room room = null;
                int startNumber = 0;

                foreach (Level level in levels)
                {
                    PlanTopology planTopology = doc.get_PlanTopology(level);
                    foreach (PlanCircuit pc in planTopology.Circuits)
                    {
                        if (!pc.IsRoomLocated)
                        {
                            room = doc.Create.NewRoom(null, pc);
                            AutoTagRooms(doc, level, roomTag);
                        }
                    }
                }

                List<Room> rooms = GetRooms(doc);

                foreach (Room r in rooms)
                {
                    // set the Room Number and Name
                    int newRoomNumber = startNumber++;
                    string newRoomName = "Room " + newRoomNumber;
                    r.Name = newRoomName;
                    r.Number = Convert.ToString(newRoomNumber);
                }
                transaction.Commit();
            }
            TaskDialog.Show("Revit", "Rooms have been successfully placed.");
            return Result.Succeeded;
        }
        private RoomTagType GetRoomTag(Document doc)
        {
            var roomTag = new FilteredElementCollector(doc)
            .OfClass(typeof(FamilySymbol))
            .OfCategory(BuiltInCategory.OST_RoomTags)
            .Cast<RoomTagType>()
            .Where(x => x.Name.Equals("Марки помещений"))
            .FirstOrDefault();

            return roomTag;
        }
        private List<Level> GetLevels(Document doc)
        {
            var levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .OfType<Level>()
                .ToList();
            return levels;
        }
        private List<Room> GetRooms(Document doc)
        {
            var rooms = new FilteredElementCollector(doc)
                .OfClass(typeof(SpatialElement)) //parent class of Room
                .WhereElementIsNotElementType()
                .Where(room => room.GetType() == typeof(Room))
                .ToList();

            return new List<Room>(rooms.Select(r => r as Room));
        }
        public void AutoTagRooms(Document doc, Level level, RoomTagType tagType)
        {
            PlanTopology planTopology = doc.get_PlanTopology(level);

            foreach (ElementId eid in planTopology.GetRoomIds())
            {
                Room tmpRoom = doc.GetElement(eid) as Room;

                if (doc.GetElement(tmpRoom.LevelId) != null && tmpRoom.Location != null)
                {
                    
                   LocationPoint locationPoint = tmpRoom.Location as LocationPoint;
                    UV point = new UV(locationPoint.Point.X, locationPoint.Point.Y);
                    RoomTag newTag = doc.Create.NewRoomTag(new LinkElementId(tmpRoom.Id), point, null);
                    newTag.RoomTagType = tagType;
                }
            }

        }
    }
}
