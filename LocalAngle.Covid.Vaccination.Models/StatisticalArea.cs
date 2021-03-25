namespace LocalAngle.Covid.Vaccination.Models
{
    /// <summary>
    /// Could represent a Middle Layer Super Output Area, Local Authority or even a region
    /// </summary>
    public class StatisticalArea
    {
        public string Code { get; set; }

        public string Name { get; set; }

        public double PopulationUnder16 { get; set; }

        public double Population16To49 { get; set; }

        public double Population50To54 { get; set; }

        public double Population55To59 { get; set; }

        public double Population60To64 { get; set; }

        public double Population65To69 { get; set; }

        public double Population70To74 { get; set; }

        public double Population75To79 { get; set; }

        public double PopulationOver80 { get; set; }

        public double PopulationOverall
        {
            get
            {
                return Population16To49 + Population50To54 + Population55To59 + Population60To64 + Population65To69 + Population70To74 + Population75To79 + PopulationOver80;
            }
        }
    }
}