using System.ComponentModel.DataAnnotations;

namespace NetAssessment.Models
{
    public class AssignmentList
    {
        public int Id { get; set; }

        [Required]
        public string Task { get; set; }

        [Required]
        public string Customer { get; set; }

        [Required]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public int Mandays { get; set; }

        [Required]
        [DataType(DataType.Text)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        [Display(Name = "Start Date")]
        public DateOnly StartDate { get; set; }

        [Display(Name = "Add Days")]
        [DisplayFormat(DataFormatString = "{0:N0}")]
        public int? AddDays { get; set; }

        [Required]
        [DataType(DataType.Text)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        [Display(Name = "Estimated Completion Date")]
        public DateOnly EstimatedCompletionDate { get; set; }

        [Required]
        public string PIC { get; set; }

        [Required]
        public string Status { get; set; }

        [DataType(DataType.Text)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        [Display(Name = "Completion Date")]
        public DateOnly? CompletionDate { get; set; }

        public string? Remarks { get; set; }

        [Required]
        [DataType(DataType.Text)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        [Display(Name = "Last Communication Date")]
        public DateOnly LastCommunicationDate { get; set; }
    }
}
