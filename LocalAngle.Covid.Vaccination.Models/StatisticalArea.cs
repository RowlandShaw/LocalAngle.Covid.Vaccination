namespace LocalAngle.Covid.Vaccination.Models
{
    /// <summary>
    /// Could represent a Middle Layer Super Output Area, Local Authority or even a region
    /// </summary>
    public class StatisticalArea
    {
        public string Code { get; set; }

        public string Name { get; set; }

        public int PopulationUnder16 { get; set; }

        public int Population16To59 { get; set; }

        public int Population60To64 { get; set; }

        public int Population65To69 { get; set; }

        public int Population70To74 { get; set; }

        public int Population75To79 { get; set; }

        public int PopulationOver80 { get; set; }
    }
}