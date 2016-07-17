namespace BUR_UI.Entities
{
    public class Items
    {
        public string Classification { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public float Amount { get; set; }
        public string BUR_Number { get; set; }

        public BURModel BUR { get; set; }
    }
}