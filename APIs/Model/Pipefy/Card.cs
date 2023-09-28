//using Microsoft.Office.Interop.Excel;
//using RpaLib.APIs.Model.Pipefy.Legacy;
using System;
using System.Collections.Generic;
using System.Linq;
//using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Model.Pipefy
{
    using Id = String;
    public class Card
    {
        public int Age { get; set; }
        //public Users Assignees{ get; set; }
        //public Attachments Attachments{ get; set; }
        public int AttachmentsCount { get; set; }
        //public CardAssignees CardAssignees { get; set; }
        public int ChecklistItemsCheckedCount { get; set; }
        public int ChecklistItemsCount { get; set; }
        //public CardRelationships ChildRelations { get; set; }
        //public Comments Comments { get; set; }
        public int CommentsCount { get; set; }
        public DateTime CreatedAt { get; set; }
        //public User CreatedBy { get; set; }
        public string CreatorEmail { get; set; }
        //public CardLateness CurrentLateness { get; set; }
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
        //public InboxEmails InboxEmails { get; set; }
        //public Labels Labels { get; set; }
        public bool Late { get; set; }
        public bool Overdue { get; set; }
        //public CardRelationships ParentRelations { get; set; }
        public string Path { get; set; }
        //public PhaseDetails PhasesHistory { get; set; }
        //public Pipe Pipe { get; set; }
        public string PublicFormSubmitterEmail { get; set; }
        public DateTime StartedCurrentPhaseAt { get; set; }
        public List<Card> Subtitles { get; set; }
        public string Suid { get; set; }
        //public Summaries Summary { get; set; }
        //public Summaries SummaryAttributes { get; set; }
        //public Summaries SummaryFields { get; set; }
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
