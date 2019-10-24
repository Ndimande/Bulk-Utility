
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

 namespace RESTAPITest1

{
    // collection + links + results
    public class Poco
    {
        // The results section of the JSON is a List - so we save the data in a list of PocoResut
        private List<PocoResult> results;
        //public List<PocoResult> Results { get; set; }
        public List<PocoResult> Results { get => results; set => results = value; }
    }

    // Result => Data + Links + Search_result_metadata 
    public class PocoResult
    {
        public PocoDataContianer Data { get; set; }
    }

    // Result => Data => properties + version
    public class PocoDataContianer
    {
        public PocoDataPropertiesContainer properties { get; set; }
        public PocoDataVersionContainer versions { get; set; }
    }
    //obj.Data.properties.Id
    // Result => Data => Properties => ...
    public class PocoDataPropertiesContainer
    {
        public string Create_date { get; set; }
        public int Create_user_id { get; set; }
        public int Id { get; set; }
        public string Name { get; set; }
        public string Mime_type { get; set; }
        public string Modify_date { get; set; }
        public int Parent_id { get; set; }
        public string Owner { get; set; }
        public int Owner_user_id { get; set; }
        public string Type_name { get; set; }







        public int Type { get; set; }
        public string volume_id { get; set; }
        public int user_id { get; set; }
        public string description { get; set; }
        public string create_date { get; set; }
        public string reserved { get; set; }
        public int modify_user_id { get; set; }
        public bool reserved_user_id { get; set; }
        public string reserved_date { get; set; }
        public string icon { get; set; }
        public string original_id { get; set; }
        public string container { get; set; }
        public string size { get; set; }
        public string perm_see { get; set; }
        public string perm_see_contents { get; set; }
        public string perm_modify { get; set; }
        public string perm_modify_attributes { get; set; }
        public string perm_modify_permissions { get; set; }
        public bool perm_create { get; set; }
        public bool perm_delete { get; set; }
        public bool perm_delete_versions { get; set; }
        public bool perm_reserve { get; set; }
        public bool perm_add_major_version { get; set; }
      
      
    }

    // Result => Data => Versions => ...
    public class PocoDataVersionContainer
    {
        public string File_name { get; set; }
        public int File_size { get; set; }
        public string File_type { get; set; }
    }

}
