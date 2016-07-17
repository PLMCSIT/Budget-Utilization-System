namespace BUR_UI.Entities
{
    public class UserModel
    {
        public string User_Number { get; set; }
        public string User_Name { get; set; }
        public string User_Password { get; set; }
        public string Discriminator { get; set; }
        public string Position { get; set; }
        public string Picture { get; set; }
    }
}
