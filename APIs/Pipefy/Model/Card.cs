//using Microsoft.Office.Interop.Excel;
//using RpaLib.APIs.Model.Pipefy.Legacy;
using System;
using System.Collections.Generic;
using System.Linq;
//using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy.Model
{
    using Id = String;
    public class Card
    {
        public int Age { get; set; }
        public List<User> Assignees{ get; set; }
        //public List<Attachment> Attachments{ get; set; }
        public int AttachmentsCount { get; set; }
        //public List<CardAssignee> Cardassignees { get; set; }
        public int ChecklistItemsCheckedCount { get; set; }
        public int ChecklistItemsCount { get; set; }
        //public List<CardRelationship> ChildRelations { get; set; }
        //public List<Comment> Comments { get; set; }
        public int CommentsCount { get; set; }
        public DateTime Createdat { get; set; }
        public User Createdby { get; set; }
        public string Creatoremail { get; set; }
        //public CardLateness Currentlateness { get; set; }
        public Phase CurrentPhase { get; set; }
        public int CurrentPhaseAge { get; set; }
        public bool Done { get; set; }
        public DateTime DueDate { get; set; }
        public string EmailMessagingAddress { get; set; }
        //public CardExpiration expiration { get; set; }
        public bool Expired { get; set; }
        public List<CardField> Fields { get; set; }
        public DateTime FinishedAt { get; set; }
        public Id Id { get; set; }
        //public List<InboxEmail> InboxEmails { get; set; }
        //public List<Label> Labels { get; set; }
        public bool Late { get; set; }
        public bool Overdue { get; set; }
        //public List<CardRelationship> ParentRelations { get; set; }
        public string Path { get; set; }
        //public List<PhaseDetail> PhasesHistory { get; set; }
        //public Pipe Pipe { get; set; }
        public string PublicFormSubmitterEmail { get; set; }
        public DateTime StartedCurrentPhaseAt { get; set; }
        public List<Card> Subtitles { get; set; }
        public string Suid { get; set; }
        //public List<Summary> Summary { get; set; }
        //public List<Summary> SummaryAttributes { get; set; }
        //public List<Summary> SummaryFields { get; set; }
        public string Title { get; set; }
        public DateTime UpdatedAt { get; set; }
        public string Url { get; set; }
        public string Uuid { get; set; }

        /* 
         * Deprecated:
         * CreatedAt
         * CreatedBy
         * Repo
         */

    }
}
