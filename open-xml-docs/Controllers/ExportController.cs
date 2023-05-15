using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace open_xml_docs.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExportController : ControllerBase
    {

        [HttpPost("export-reportToMSWord/{id}/{projectId}/{userId}")]
        public async Task<IActionResult> GetReportToMSWord(int id, int projectId, int userId, [FromBody] GetUserPostDTO aDto)
        {
            try
            {

                var uploadFolder = Path.Combine(_hostingEnvironment.WebRootPath, "WordFile");
                if (!Directory.Exists(uploadFolder))
                {
                    Directory.CreateDirectory(uploadFolder);
                }
                string uniqueFileName = Guid.NewGuid().ToString() + ".docx";
                string filepath = Path.Combine(uploadFolder, uniqueFileName);

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create
                    (filepath, WordprocessingDocumentType.Document))
                {

                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                    // Create the document structure and add some text.
                    mainPart.Document = new Document();



                    Body body = mainPart.Document.AppendChild(new Body());
                    Paragraph paraTitle = body.AppendChild(new Paragraph());
                    Run runTitle = paraTitle.AppendChild(new Run());
                    runTitle.AppendChild(new Text() { Text = "Meeting Note Title: ", Space = SpaceProcessingModeValues.Preserve });


                    RunProperties runTitlePro = new RunProperties();
                    runTitlePro.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    runTitle.RunProperties = runTitlePro;

                    Run runTitle1 = paraTitle.AppendChild(new Run());
                    runTitle1.AppendChild(new Text() { Text = reportData.MeetingTitle });


                    Paragraph paraMeetingNumber = body.AppendChild(new Paragraph());
                    Run runMeetingNumber = paraMeetingNumber.AppendChild(new Run());
                    runMeetingNumber.AppendChild(new Text() { Text = "Meeting Number: ", Space = SpaceProcessingModeValues.Preserve });

                    RunProperties runMeetingNumberPro = new RunProperties();
                    runMeetingNumberPro.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    runMeetingNumber.RunProperties = runMeetingNumberPro;


                    Run paraMeetingNumber1 = paraMeetingNumber.AppendChild(new Run());
                    paraMeetingNumber1.AppendChild(new Text() { Text = reportData.MeetingNumber.ToString() });


                    Paragraph paraLocation = body.AppendChild(new Paragraph());
                    Run runLocation = paraLocation.AppendChild(new Run());
                    runLocation.AppendChild(new Text() { Text = "Location: ", Space = SpaceProcessingModeValues.Preserve });

                    RunProperties runLocationPro = new RunProperties();
                    runLocationPro.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    runLocation.RunProperties = runLocationPro;

                    Run paraLocationvalue = paraLocation.AppendChild(new Run());
                    paraLocationvalue.AppendChild(new Text() { Text = reportData.Location.ToString() });

                    Paragraph paraMeetingStatus = body.AppendChild(new Paragraph());
                    Run runMeetingStatus = paraMeetingStatus.AppendChild(new Run());
                    runMeetingStatus.AppendChild(new Text() { Text = "Status: ", Space = SpaceProcessingModeValues.Preserve });

                    RunProperties runMeetingStatusPro = new RunProperties();
                    runMeetingStatusPro.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    runMeetingStatus.RunProperties = runMeetingStatusPro;

                    var _status = reportData.Status == (int)MeetingStatus.Draft ?
                                                                   MeetingStatus.Draft : reportData.Status == (int)MeetingStatus.InReview
                                                                   ? MeetingStatus.InReview : MeetingStatus.Finalized;

                    Run paraMeetingStatusvalue = paraMeetingStatus.AppendChild(new Run());
                    paraMeetingStatusvalue.AppendChild(new Text() { Text = _status.ToString() });

                    Paragraph paraMeetingDate = body.AppendChild(new Paragraph());
                    Run runMeetingDate = paraMeetingDate.AppendChild(new Run());
                    runMeetingDate.AppendChild(new Text() { Text = "Meeting Date: ", Space = SpaceProcessingModeValues.Preserve });

                    RunProperties runMeetingDatePro = new RunProperties();
                    runMeetingDatePro.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    runMeetingDate.RunProperties = runMeetingDatePro;

                    Run paraMeetingDatevalue = paraMeetingDate.AppendChild(new Run());
                    paraMeetingDatevalue.AppendChild(new Text() { Text = reportData.MeetingDate == null ? "" : Convert.ToDateTime(reportData.MeetingDate).ToString("dd-MM-yyyy") });

                    Paragraph paraProjectInfo = body.AppendChild(new Paragraph());
                    Run runProjectInfo = paraProjectInfo.AppendChild(new Run());
                    runProjectInfo.AppendChild(new Text("General Information: "));

                    RunProperties run1Properties = new RunProperties();
                    run1Properties.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    runProjectInfo.RunProperties = run1Properties;

                    Paragraph paraProjectTitle = body.AppendChild(new Paragraph());
                    Run runProjectTitle = paraProjectTitle.AppendChild(new Run());
                    runProjectTitle.AppendChild(new Text() { Text = "Project Title: ", Space = SpaceProcessingModeValues.Preserve });

                    RunProperties runProjectTitlePro = new RunProperties();
                    runProjectTitlePro.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    runProjectTitle.RunProperties = runProjectTitlePro;

                    Run paraProjectTitlevalue = paraProjectTitle.AppendChild(new Run());
                    paraProjectTitlevalue.AppendChild(new Text() { Text = projectDetails.title });

                    Paragraph paraProjectContactNumber = body.AppendChild(new Paragraph());
                    Run runProjectContactNumber = paraProjectContactNumber.AppendChild(new Run());
                    runProjectContactNumber.AppendChild(new Text() { Text = "Contact Number: ", Space = SpaceProcessingModeValues.Preserve });

                    RunProperties runProjectContactNumberPro = new RunProperties();
                    runProjectContactNumberPro.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    runProjectContactNumber.RunProperties = runProjectContactNumberPro;

                    Run paraProjectContactNumbervalue = paraProjectContactNumber.AppendChild(new Run());
                    paraProjectContactNumbervalue.AppendChild(new Text() { Text = projectDetails.contract_number });

                    Paragraph paraProjectClient = body.AppendChild(new Paragraph());
                    Run runProjectClient = paraProjectClient.AppendChild(new Run());
                    runProjectClient.AppendChild(new Text() { Text = "Client: ", Space = SpaceProcessingModeValues.Preserve });

                    RunProperties runProjectClientPro = new RunProperties();
                    runProjectClientPro.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    runProjectClient.RunProperties = runProjectClientPro;

                    Run runProjectClientvalue = paraProjectClient.AppendChild(new Run());
                    runProjectClientvalue.AppendChild(new Text() { Text = projectDetails.client });

                    Paragraph paraProjectStartingDate = body.AppendChild(new Paragraph());
                    Run runProjectStartingDate = paraProjectStartingDate.AppendChild(new Run());
                    runProjectStartingDate.AppendChild(new Text() { Text = "Project Starting Date: ", Space = SpaceProcessingModeValues.Preserve });

                    RunProperties runProjectStartingDatePro = new RunProperties();
                    runProjectStartingDatePro.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    runProjectStartingDate.RunProperties = runProjectStartingDatePro;

                    Run paraProjectStartingDatevalue = paraProjectStartingDate.AppendChild(new Run());
                    paraProjectStartingDatevalue.AppendChild(new Text() { Text = Convert.ToDateTime(projectDetails.project_starting_date).ToString("dd-MM-yyyy") });

                    Paragraph paraProjectCompletionDate = body.AppendChild(new Paragraph());
                    Run runProjectCompletionDate = paraProjectCompletionDate.AppendChild(new Run());
                    runProjectCompletionDate.AppendChild(new Text() { Text = "Project Expected Completion Date: ", Space = SpaceProcessingModeValues.Preserve });

                    RunProperties runProjectCompletionDatePro = new RunProperties();
                    runProjectCompletionDatePro.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    runProjectCompletionDate.RunProperties = runProjectCompletionDatePro;

                    Run runProjectCompletionDatevalue = paraProjectCompletionDate.AppendChild(new Run());
                    runProjectCompletionDatevalue.AppendChild(new Text() { Text = Convert.ToDateTime(projectDetails.exp_completion_datetime).ToString("dd-MM-yyyy") });

                    ///Attendees
                    Paragraph paraAttendees = body.AppendChild(new Paragraph());
                    Run runAttendees = paraAttendees.AppendChild(new Run());
                    runAttendees.AppendChild(new Text("Attendees: "));

                    RunProperties runAttendesProp = new RunProperties();
                    runAttendesProp.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    runAttendees.RunProperties = runAttendesProp;

                    foreach (var group in reportData.MeetingGroups)
                    {
                        Paragraph paraGroups = body.AppendChild(new Paragraph());
                        Run runGroups = paraGroups.AppendChild(new Run());
                        runGroups.AppendChild(new Text(group.Name));

                        RunProperties runGroupsProp = new RunProperties();
                        runGroupsProp.Append(new Bold());
                        //set the first runs RunProperties to the RunProperties containing the bold

                        runGroups.RunProperties = runGroupsProp;


                        Table table = new Table();

                        TableProperties props = new TableProperties(
                            new TableCellMarginDefault(
                                new TopMargin() { Width = "50", Type = TableWidthUnitValues.Dxa },
                                new StartMargin() { Width = "100", Type = TableWidthUnitValues.Dxa },
                                new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                                new EndMargin() { Width = "100", Type = TableWidthUnitValues.Dxa }),

                        new TableBorders(
                            new TopBorder
                            {
                                Val = new EnumValue<BorderValues>(BorderValues.Single),
                                Size = 4
                            },
                            new BottomBorder
                            {
                                Val = new EnumValue<BorderValues>(BorderValues.Single),
                                Size = 4
                            },
                            new LeftBorder
                            {
                                Val = new EnumValue<BorderValues>(BorderValues.Single),
                                Size = 4
                            },
                            new RightBorder
                            {
                                Val = new EnumValue<BorderValues>(BorderValues.Single),
                                Size = 4
                            },
                            new InsideHorizontalBorder
                            {
                                Val = new EnumValue<BorderValues>(BorderValues.Single),
                                Size = 4
                            },
                            new InsideVerticalBorder
                            {
                                Val = new EnumValue<BorderValues>(BorderValues.Single),
                                Size = 4
                            }));

                        table.AppendChild<TableProperties>(props);

                        var title = new TableRow();

                        Run tableAttendeesId = new Run();
                        tableAttendeesId.AppendChild(new Text("ID"));

                        RunProperties tableAttendessIdProp = new RunProperties();
                        tableAttendessIdProp.Append(new Bold());
                        //set the first runs RunProperties to the RunProperties containing the bold
                        tableAttendeesId.RunProperties = tableAttendessIdProp;

                        Run tableAttendessName = new Run();
                        tableAttendessName.AppendChild(new Text("Name"));

                        RunProperties tableAttendessNameProp = new RunProperties();
                        tableAttendessNameProp.Append(new Bold());
                        //set the first runs RunProperties to the RunProperties containing the bold
                        tableAttendessName.RunProperties = tableAttendessNameProp;

                        Run tableAttendesEmail = new Run();
                        tableAttendesEmail.AppendChild(new Text("Email"));

                        RunProperties tableAttendessEmailProp = new RunProperties();
                        tableAttendessEmailProp.Append(new Bold());
                        //set the first runs RunProperties to the RunProperties containing the bold
                        tableAttendesEmail.RunProperties = tableAttendessEmailProp;


                        Run tableAttendesIsAbsent = new Run();
                        tableAttendesIsAbsent.AppendChild(new Text("Is Absent"));

                        RunProperties tableAttendessIsAbsentProp = new RunProperties();
                        tableAttendessIsAbsentProp.Append(new Bold());
                        //set the first runs RunProperties to the RunProperties containing the bold
                        tableAttendesIsAbsent.RunProperties = tableAttendessIsAbsentProp;

                        TableCell tc00 = new TableCell(new Paragraph(tableAttendeesId));
                        TableCell tc01 = new TableCell(new Paragraph(tableAttendessName));
                        TableCell tc02 = new TableCell(new Paragraph(tableAttendesEmail));
                        TableCell tc03 = new TableCell(new Paragraph(tableAttendesIsAbsent));

                        tc03.Append(new TableCellProperties(
                                       new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2000" }));

                        title.Append(tc00, tc01, tc02, tc03);
                        table.Append(title);

                        foreach (var user in group.MeetingProjectUsers)
                        {
                            var tr = new TableRow();

                            // Add a cell to each column in the row.
                            TableCell tc1 = new TableCell(new Paragraph(new Run(new Text(user.UserId.ToString()))));
                            TableCell tc2 = new TableCell(new Paragraph(new Run(new Text(user.UserName))));
                            TableCell tc3 = new TableCell(new Paragraph(new Run(new Text(user.Email))));
                            TableCell tc4 = new TableCell(new Paragraph(new Run(new Text(user.IsAbsent == true ? "True" : "False"))));

                            tr.Append(tc1, tc2, tc3, tc4);

                            table.Append(tr);
                        }

                        foreach (var user in group.MeetingExternalUsers)
                        {
                            var tr = new TableRow();

                            // Add a cell to each column in the row.
                            TableCell tc1 = new TableCell(new Paragraph(new Run(new Text(user.Id.ToString()))));
                            TableCell tc2 = new TableCell(new Paragraph(new Run(new Text(user.Name))));
                            TableCell tc3 = new TableCell(new Paragraph(new Run(new Text(user.Email))));
                            TableCell tc4 = new TableCell(new Paragraph(new Run(new Text(user.IsAbsent == true ? "True" : "False"))));

                            tr.Append(tc1, tc2, tc3, tc4);

                            table.Append(tr);
                        }

                        mainPart.Document.Body.Append(table);

                        Paragraph paraGroupsBreak = mainPart.Document.Body.AppendChild(new Paragraph());
                        Run runGroupsBreak = paraGroupsBreak.AppendChild(new Run());
                        //runGroupsBreak.AppendChild(new Break());

                    }

                    Paragraph paraTopic = body.AppendChild(new Paragraph());
                    Run runTopic = paraTopic.AppendChild(new Run());
                    runTopic.AppendChild(new Break());
                    runTopic.AppendChild(new Text("Topic: "));

                    RunProperties runTopicProp = new RunProperties();
                    runTopicProp.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    runTopic.RunProperties = runTopicProp;

                    foreach (var topic in reportData.MeetingTopics)
                    {
                        topic.TopicPriority = topic.TopicPriority + 1;

                        Paragraph paraTopicDetails = body.AppendChild(new Paragraph());
                        Run runTopicDetails = paraTopicDetails.AppendChild(new Run());
                        runTopicDetails.AppendChild(new Text(topic.TopicPriority + " " + topic.Name));

                        foreach (var subtopic in topic.MeetingSubTopics)
                        {
                            subtopic.SubTopicPriority = subtopic.SubTopicPriority + 1;

                            Paragraph paraSubTopic = body.AppendChild(new Paragraph());
                            Run runSubTopic = paraSubTopic.AppendChild(new Run());
                            runSubTopic.AppendChild(new TabChar());
                            runSubTopic.AppendChild(new Text(topic.TopicPriority + "." + subtopic.SubTopicPriority + " " + subtopic.Name));

                        }
                    }


                    Paragraph paraTags = body.AppendChild(new Paragraph());
                    Run runTags = paraTags.AppendChild(new Run());
                    runTags.AppendChild(new Break());
                    runTags.AppendChild(new Text("Tags: "));

                    RunProperties runTagsProp = new RunProperties();
                    runTagsProp.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold

                    runTags.RunProperties = runTagsProp;

                    var we = reportData.MeetingTags.Select(tags => tags.Name.ToString());

                    var _tags = string.Join(", ", reportData.MeetingTags.Select(tags => tags.Name.ToString()).ToList());

                    Paragraph paraTagsDetails = body.AppendChild(new Paragraph());
                    Run runTagDetails = paraTagsDetails.AppendChild(new Run());
                    runTagDetails.AppendChild(new Text(_tags));
                    runTagDetails.AppendChild(new Break());
                    runTagDetails.AppendChild(new Break());

                    Paragraph paraMeetingNotes = body.AppendChild(new Paragraph());
                    Run runMeetingNotes = paraMeetingNotes.AppendChild(new Run());
                    runMeetingNotes.AppendChild(new Text("Meeting Notes: "));

                    RunProperties runMeetingNotesProp = new RunProperties();
                    runMeetingNotesProp.Append(new Bold());
                    runMeetingNotesProp.Append(new Underline() { Val = UnderlineValues.Single });

                    FontSize fontSize = new FontSize();
                    fontSize.Val = "36";

                    runMeetingNotesProp.Append(fontSize);

                    //set the first runs RunProperties to the RunProperties containing the bold

                    runMeetingNotes.RunProperties = runMeetingNotesProp;

                    foreach (var topic in reportData.MeetingTopics)
                    {
                        Paragraph paraTopicInDetails = body.AppendChild(new Paragraph());
                        Run runTopicInDetails = paraTopicInDetails.AppendChild(new Run());
                        runTopicInDetails.AppendChild(new Break());
                        runTopicInDetails.AppendChild(new Text(topic.TopicPriority + " " + topic.Name));

                        RunProperties runTopicInDetailsProp = new RunProperties();

                        FontSize fontSize36 = new FontSize();
                        fontSize36.Val = "36";
                        runTopicInDetailsProp.Append(fontSize36);

                        runTopicInDetails.RunProperties = runTopicInDetailsProp;

                        foreach (var subtopic in topic.MeetingSubTopics)
                        {
                            Paragraph paraSubTopicDetails = body.AppendChild(new Paragraph());
                            Run runSubTopicInDetails = paraSubTopicDetails.AppendChild(new Run());
                            //runSubTopicInDetails.AppendChild(new TabChar());
                            runSubTopicInDetails.AppendChild(new Text(topic.TopicPriority + "." + subtopic.SubTopicPriority + " " + subtopic.Name));

                            RunProperties runSubTopicInDetailsProp = new RunProperties();

                            FontSize fontSize14 = new FontSize();
                            fontSize14.Val = "28";

                            runSubTopicInDetailsProp.Append(fontSize14);
                            runSubTopicInDetails.RunProperties = runSubTopicInDetailsProp;

                            foreach (var notes in subtopic.MeetingNotes)
                            {
                                Paragraph paraNotes = body.AppendChild(new Paragraph());
                                Run runNotes = paraNotes.AppendChild(new Run());
                                //runNotes.AppendChild(new TabChar());
                                //runNotes.AppendChild(new TabChar());

                                runNotes.AppendChild(new Text(topic.TopicPriority + "." + subtopic.SubTopicPriority + "." + notes.NoteNumber + " " + notes.Name));



                                Paragraph meetingNotes_Note = body.AppendChild(new Paragraph());
                                Run runMeetingNote_Note = meetingNotes_Note.AppendChild(new Run());
                                runMeetingNote_Note.AppendChild(new Text() { Text = "Note: ", Space = SpaceProcessingModeValues.Preserve });


                                RunProperties runMeetingNote_Pro = new RunProperties();
                                runMeetingNote_Pro.Append(new Bold());
                                //set the first runs RunProperties to the RunProperties containing the bold
                                runMeetingNote_Note.RunProperties = runMeetingNote_Pro;

                                Run runMeeting_Note_Value = meetingNotes_Note.AppendChild(new Run());
                                runMeeting_Note_Value.AppendChild(new Text() { Text = notes.Note });



                                Paragraph paraNoteTaskTitle = body.AppendChild(new Paragraph());
                                Run runNoteTaskTitle = paraNoteTaskTitle.AppendChild(new Run());
                                runNoteTaskTitle.AppendChild(new Break());
                                runNoteTaskTitle.AppendChild(new Text("Tasks"));

                                RunProperties runMeetingNotesTaskTitleProp = new RunProperties();
                                runMeetingNotesTaskTitleProp.Append(new Bold());
                                //set the first runs RunProperties to the RunProperties containing the bold
                                runNoteTaskTitle.RunProperties = runMeetingNotesTaskTitleProp;


                                Table notesTasktable = new Table();


                                TableProperties propsNotesTask = new TableProperties(
                                   new TableCellMarginDefault(
                                       new TopMargin() { Width = "50", Type = TableWidthUnitValues.Dxa },
                                       new StartMargin() { Width = "100", Type = TableWidthUnitValues.Dxa },
                                       new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                                       new EndMargin() { Width = "100", Type = TableWidthUnitValues.Dxa }),

                               new TableBorders(
                                   new TopBorder
                                   {
                                       Val = new EnumValue<BorderValues>(BorderValues.Single),
                                       Size = 4
                                   },
                                   new BottomBorder
                                   {
                                       Val = new EnumValue<BorderValues>(BorderValues.Single),
                                       Size = 4
                                   },
                                   new LeftBorder
                                   {
                                       Val = new EnumValue<BorderValues>(BorderValues.Single),
                                       Size = 4
                                   },
                                   new RightBorder
                                   {
                                       Val = new EnumValue<BorderValues>(BorderValues.Single),
                                       Size = 4
                                   },
                                   new InsideHorizontalBorder
                                   {
                                       Val = new EnumValue<BorderValues>(BorderValues.Single),
                                       Size = 4
                                   },
                                   new InsideVerticalBorder
                                   {
                                       Val = new EnumValue<BorderValues>(BorderValues.Single),
                                       Size = 4
                                   }));



                                notesTasktable.AppendChild<TableProperties>(propsNotesTask);


                                if (notes.MeetingNoteTasks.Count > 0)
                                {
                                    var noteTaskstitle = new TableRow();


                                    Run tableID = new Run();
                                    tableID.AppendChild(new Text("ID"));

                                    RunProperties tableIdProp = new RunProperties();
                                    tableIdProp.Append(new Bold());
                                    //set the first runs RunProperties to the RunProperties containing the bold
                                    tableID.RunProperties = tableIdProp;

                                    Run tableName = new Run();
                                    tableName.AppendChild(new Text("Task Name"));

                                    RunProperties tableNameProp = new RunProperties();
                                    tableNameProp.Append(new Bold());
                                    //set the first runs RunProperties to the RunProperties containing the bold
                                    tableName.RunProperties = tableNameProp;

                                    Run tableMeetingNoteCreatedTaskOnMeetingBoard = new Run();
                                    tableMeetingNoteCreatedTaskOnMeetingBoard.AppendChild(new Text("Created Task in Meeting Board"));

                                    RunProperties tableMeetingNoteCreatedTaskOnMeetingBoardProp = new RunProperties();
                                    tableMeetingNoteCreatedTaskOnMeetingBoardProp.Append(new Bold());
                                    //set the first runs RunProperties to the RunProperties containing the bold
                                    tableMeetingNoteCreatedTaskOnMeetingBoard.RunProperties = tableMeetingNoteCreatedTaskOnMeetingBoardProp;


                                    Run tableAssignTo = new Run();
                                    tableAssignTo.AppendChild(new Text("Assigned To"));

                                    RunProperties tableAssignToProp = new RunProperties();
                                    tableAssignToProp.Append(new Bold());
                                    //set the first runs RunProperties to the RunProperties containing the bold
                                    tableAssignTo.RunProperties = tableAssignToProp;

                                    Run tableDueDate = new Run();
                                    tableDueDate.AppendChild(new Text("Due Date"));

                                    RunProperties tableDueDateProp = new RunProperties();
                                    tableDueDateProp.Append(new Bold());
                                    //set the first runs RunProperties to the RunProperties containing the bold
                                    tableDueDate.RunProperties = tableDueDateProp;


                                    TableCell ntc000 = new TableCell(new Paragraph(tableID));
                                    TableCell ntc001 = new TableCell(new Paragraph(tableName));
                                    TableCell ntc002 = new TableCell(new Paragraph(tableMeetingNoteCreatedTaskOnMeetingBoard));
                                    TableCell ntc003 = new TableCell(new Paragraph(tableAssignTo));
                                    TableCell ntc004 = new TableCell(new Paragraph(tableDueDate));


                                    noteTaskstitle.Append(ntc000, ntc001, ntc002, ntc003, ntc004);
                                    notesTasktable.Append(noteTaskstitle);
                                }


                                foreach (var tasks in notes.MeetingNoteTasks)
                                {
                                    //Paragraph paraTasks = body.AppendChild(new Paragraph());
                                    //Run runTasks = paraTasks.AppendChild(new Run());
                                    //runTasks.AppendChild(new Text(topic.TopicNumber + "." + subtopic.SubTopicNumber + "." + notes.NoteNumber +" " + no.Name));

                                    var tr1 = new TableRow();


                                    var _date = Convert.ToDateTime(tasks.StartDate).ToString("dd-MM-yyyy") + " to " + Convert.ToDateTime(tasks.EndDate).ToString("dd-MM-yyyy");

                                    _date = tasks.StartDate == null ? "" : _date;

                                    // Add a cell to each column in the row.
                                    TableCell tc1 = new TableCell(new Paragraph(new Run(new Text(tasks.Id.ToString()))));
                                    TableCell tc2 = new TableCell(new Paragraph(new Run(new Text(tasks.Name))));
                                    TableCell tc3 = new TableCell(new Paragraph(new Run(new Text(tasks.IsWorkBoard == true ? "Yes" : "No"))));
                                    TableCell tc4 = new TableCell(new Paragraph(new Run(new Text(tasks.AssignedTo))));
                                    TableCell tc5 = new TableCell(new Paragraph(new Run(new Text(_date.ToString()))));



                                    // Assume you want columns that are automatically sized.
                                    tc1.Append(new TableCellProperties(
                                        new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                                    tc2.Append(new TableCellProperties(
                                       new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }));

                                    tc3.Append(new TableCellProperties(
                                       new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                                    tc4.Append(new TableCellProperties(
                                      new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "4500" }));

                                    tc5.Append(new TableCellProperties(
                                       new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "4000" }));

                                    tr1.Append(tc1, tc2, tc3, tc4, tc5);

                                    notesTasktable.Append(tr1);


                                }

                                mainPart.Document.Body.Append(notesTasktable);


                                Paragraph paraNoteFiles = body.AppendChild(new Paragraph());
                                Run runNoteFiles = paraNoteFiles.AppendChild(new Run());
                                runNoteFiles.AppendChild(new Break());
                                runNoteFiles.AppendChild(new Text("Files"));

                                RunProperties runMeetingNotesFilesProp = new RunProperties();
                                runMeetingNotesFilesProp.Append(new Bold());
                                //set the first runs RunProperties to the RunProperties containing the bold
                                runNoteFiles.RunProperties = runMeetingNotesFilesProp;


                                int fileIndex = 1;
                                foreach (var meetingNoteFiles in notes.MeetingNoteFiles)
                                {
                                    Paragraph paraNoteFilesDetails = body.AppendChild(new Paragraph());
                                    Run runNoteFilesDetails = paraNoteFilesDetails.AppendChild(new Run());
                                    runNoteFilesDetails.AppendChild(new Text(fileIndex + ". " + meetingNoteFiles.Name.Split("~icoPlan~")[1]));

                                    fileIndex++;
                                }


                                Paragraph paraMeetingFilesBreak = mainPart.Document.Body.AppendChild(new Paragraph());
                                Run runMeetingFilesBreak = paraMeetingFilesBreak.AppendChild(new Run());


                                Paragraph paragraphhorizontalLine = body.AppendChild(new Paragraph());

                                // Create a paragraph border with a single line at the bottom
                                ParagraphBorders borders = new ParagraphBorders(
                                    new BottomBorder() { Val = BorderValues.Double, Size = 6, Space = 1, Color = "auto" });

                                // Set the paragraph border
                                paragraphhorizontalLine.ParagraphProperties = new ParagraphProperties(borders);


                            }
                        }



                    }




                    Paragraph paraTaskFiles = body.AppendChild(new Paragraph());
                    Run runTaskFiles = paraTaskFiles.AppendChild(new Run());
                    runTaskFiles.AppendChild(new Break());
                    runTaskFiles.AppendChild(new Text("Tasks & Files: "));
                    RunProperties runTaskFilesProp = new RunProperties();
                    runTaskFilesProp.Append(new Bold());
                    runTaskFilesProp.Append(new Underline() { Val = UnderlineValues.Single });

                    FontSize fontSize142 = new FontSize();
                    fontSize142.Val = "36";

                    runTaskFilesProp.Append(fontSize142);

                    runTaskFiles.RunProperties = runTaskFilesProp;

                    Paragraph paraTaskFilesSummary = body.AppendChild(new Paragraph());
                    Run runTaskFilesSummary = paraTaskFilesSummary.AppendChild(new Run());
                    runTaskFilesSummary.AppendChild(new Text("Task Summary: "));

                    RunProperties runTaskFilesSummaryProp = new RunProperties();
                    runTaskFilesSummaryProp.Append(new Bold());

                    runTaskFilesSummary.RunProperties = runTaskFilesSummaryProp;


                    Table notestable = new Table();


                    TableProperties propsNotes = new TableProperties(
                       new TableCellMarginDefault(
                           new TopMargin() { Width = "50", Type = TableWidthUnitValues.Dxa },
                           new StartMargin() { Width = "100", Type = TableWidthUnitValues.Dxa },
                           new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                           new EndMargin() { Width = "100", Type = TableWidthUnitValues.Dxa }),

                   new TableBorders(
                       new TopBorder
                       {
                           Val = new EnumValue<BorderValues>(BorderValues.Single),
                           Size = 4
                       },
                       new BottomBorder
                       {
                           Val = new EnumValue<BorderValues>(BorderValues.Single),
                           Size = 4
                       },
                       new LeftBorder
                       {
                           Val = new EnumValue<BorderValues>(BorderValues.Single),
                           Size = 4
                       },
                       new RightBorder
                       {
                           Val = new EnumValue<BorderValues>(BorderValues.Single),
                           Size = 4
                       },
                       new InsideHorizontalBorder
                       {
                           Val = new EnumValue<BorderValues>(BorderValues.Single),
                           Size = 4
                       },
                       new InsideVerticalBorder
                       {
                           Val = new EnumValue<BorderValues>(BorderValues.Single),
                           Size = 4
                       }));



                    notestable.AppendChild<TableProperties>(propsNotes);

                    var taskstitle = new TableRow();

                    Run tableTaskId = new Run();
                    tableTaskId.AppendChild(new Text("ID"));

                    RunProperties tableTaskIdProp = new RunProperties();
                    tableTaskIdProp.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    tableTaskId.RunProperties = tableTaskIdProp;

                    Run tableTaskName = new Run();
                    tableTaskName.AppendChild(new Text("Name"));

                    RunProperties tableTaskNameProp = new RunProperties();
                    tableTaskNameProp.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    tableTaskName.RunProperties = tableTaskNameProp;


                    Run tableTaskNoteNo = new Run();
                    tableTaskNoteNo.AppendChild(new Text("Note No"));

                    RunProperties tableTaskNoteNoProp = new RunProperties();
                    tableTaskNoteNoProp.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    tableTaskNoteNo.RunProperties = tableTaskNoteNoProp;


                    Run tableTaskCreatedTaskOnMeetingBoard = new Run();
                    tableTaskCreatedTaskOnMeetingBoard.AppendChild(new Text("Created Task in Meeting Board"));

                    RunProperties tableCreatedTaskOnMeetingBoardProp = new RunProperties();
                    tableCreatedTaskOnMeetingBoardProp.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    tableTaskCreatedTaskOnMeetingBoard.RunProperties = tableCreatedTaskOnMeetingBoardProp;


                    Run tableTaskAssignedTo = new Run();
                    tableTaskAssignedTo.AppendChild(new Text("Assigned To"));

                    RunProperties tableTaskAssignedToProp = new RunProperties();
                    tableTaskAssignedToProp.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    tableTaskAssignedTo.RunProperties = tableTaskAssignedToProp;

                    Run tableTaskDueDate = new Run();
                    tableTaskDueDate.AppendChild(new Text("Due Date"));

                    RunProperties tableTaskDueDateProp = new RunProperties();
                    tableTaskDueDateProp.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    tableTaskDueDate.RunProperties = tableTaskDueDateProp;


                    TableCell ts000 = new TableCell(new Paragraph(tableTaskId));
                    TableCell ts001 = new TableCell(new Paragraph(tableTaskName));
                    TableCell ts002 = new TableCell(new Paragraph(tableTaskNoteNo));
                    TableCell ts003 = new TableCell(new Paragraph(tableTaskCreatedTaskOnMeetingBoard));
                    TableCell ts004 = new TableCell(new Paragraph(tableTaskAssignedTo));
                    TableCell ts005 = new TableCell(new Paragraph(tableTaskDueDate));




                    taskstitle.Append(ts000, ts001, ts002, ts003, ts004, ts005);
                    notestable.Append(taskstitle);



                    foreach (var topic in reportData.MeetingTopics)
                    {
                        foreach (var subtopic in topic.MeetingSubTopics)
                        {
                            foreach (var notes in subtopic.MeetingNotes)
                            {
                                foreach (var tasks in notes.MeetingNoteTasks)
                                {
                                    var tr1 = new TableRow();

                                    var _date = Convert.ToDateTime(tasks.StartDate).ToString("dd-MM-yyyy") + " to " + Convert.ToDateTime(tasks.EndDate).ToString("dd-MM-yyyy");
                                    _date = tasks.StartDate == null ? "" : _date;

                                    var noteNo = topic.TopicPriority + "." + subtopic.SubTopicPriority + "." + notes.NoteNumber;

                                    // Add a cell to each column in the row.
                                    TableCell tc1 = new TableCell(new Paragraph(new Run(new Text(tasks.Id.ToString()))));
                                    TableCell tc2 = new TableCell(new Paragraph(new Run(new Text(tasks.Name))));
                                    TableCell tc3 = new TableCell(new Paragraph(new Run(new Text(noteNo))));
                                    TableCell tc4 = new TableCell(new Paragraph(new Run(new Text(tasks.IsWorkBoard == true ? "Yes" : "No"))));
                                    TableCell tc5 = new TableCell(new Paragraph(new Run(new Text(tasks.AssignedTo))));
                                    TableCell tc6 = new TableCell(new Paragraph(new Run(new Text(_date.ToString()))));



                                    // Assume you want columns that are automatically sized.
                                    tc1.Append(new TableCellProperties(
                                        new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                                    tc2.Append(new TableCellProperties(
                                       new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" }));
                                    tc3.Append(new TableCellProperties(
                                      new TableCellWidth { Type = TableWidthUnitValues.Auto }));
                                    tc4.Append(new TableCellProperties(
                                     new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                                    tc5.Append(new TableCellProperties(
                                      new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "4500" }));

                                    tc6.Append(new TableCellProperties(
                                       new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "4500" }));

                                    tr1.Append(tc1, tc2, tc3, tc4, tc5, tc6);

                                    notestable.Append(tr1);
                                }

                            }
                        }
                    }
                    mainPart.Document.Body.Append(notestable);



                    /// Files ////////////

                    Paragraph paraFiles = body.AppendChild(new Paragraph());
                    Run runFiles = paraFiles.AppendChild(new Run());
                    runFiles.AppendChild(new Break());
                    runFiles.AppendChild(new Text("Files: "));

                    RunProperties runFilesProp = new RunProperties();
                    runFilesProp.Append(new Bold());
                    runFiles.RunProperties = runFilesProp;


                    Table filesTables = new Table();


                    TableProperties propsFiles = new TableProperties(
                       new TableCellMarginDefault(
                           new TopMargin() { Width = "50", Type = TableWidthUnitValues.Dxa },
                           new StartMargin() { Width = "100", Type = TableWidthUnitValues.Dxa },
                           new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                           new EndMargin() { Width = "100", Type = TableWidthUnitValues.Dxa }),

                   new TableBorders(
                       new TopBorder
                       {
                           Val = new EnumValue<BorderValues>(BorderValues.Single),
                           Size = 4
                       },
                       new BottomBorder
                       {
                           Val = new EnumValue<BorderValues>(BorderValues.Single),
                           Size = 4
                       },
                       new LeftBorder
                       {
                           Val = new EnumValue<BorderValues>(BorderValues.Single),
                           Size = 4
                       },
                       new RightBorder
                       {
                           Val = new EnumValue<BorderValues>(BorderValues.Single),
                           Size = 4
                       },
                       new InsideHorizontalBorder
                       {
                           Val = new EnumValue<BorderValues>(BorderValues.Single),
                           Size = 4
                       },
                       new InsideVerticalBorder
                       {
                           Val = new EnumValue<BorderValues>(BorderValues.Single),
                           Size = 4
                       }));



                    filesTables.AppendChild<TableProperties>(propsFiles);



                    var filestitle = new TableRow();


                    Run tableFileId = new Run();
                    tableFileId.AppendChild(new Text("ID"));

                    RunProperties tableFileIdProp = new RunProperties();
                    tableFileIdProp.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    tableFileId.RunProperties = tableFileIdProp;

                    Run tableFileName = new Run();
                    tableFileName.AppendChild(new Text("File Name"));

                    RunProperties tableFileNameProp = new RunProperties();
                    tableFileNameProp.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    tableFileName.RunProperties = tableFileNameProp;

                    Run tableFileNoteNo = new Run();
                    tableFileNoteNo.AppendChild(new Text("Note No"));

                    RunProperties tableNoteNoProp = new RunProperties();
                    tableNoteNoProp.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    tableFileNoteNo.RunProperties = tableNoteNoProp;


                    Run tableFileUploadedBy = new Run();
                    tableFileUploadedBy.AppendChild(new Text("Uploaded By"));

                    RunProperties tableFileUploadedByProp = new RunProperties();
                    tableFileUploadedByProp.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    tableFileUploadedBy.RunProperties = tableFileUploadedByProp;

                    Run tableFileUploadedDate = new Run();
                    tableFileUploadedDate.AppendChild(new Text("Upload Date"));

                    RunProperties tableFileUploadedDateProp = new RunProperties();
                    tableFileUploadedDateProp.Append(new Bold());
                    //set the first runs RunProperties to the RunProperties containing the bold
                    tableFileUploadedDate.RunProperties = tableFileUploadedDateProp;

                    TableCell tc000 = new TableCell(new Paragraph(tableFileId));
                    TableCell tc001 = new TableCell(new Paragraph(tableFileName));
                    TableCell tc002 = new TableCell(new Paragraph(tableFileNoteNo));
                    TableCell tc003 = new TableCell(new Paragraph(tableFileUploadedBy));
                    TableCell tc004 = new TableCell(new Paragraph(tableFileUploadedDate));



                    int noteFileCheck = 0;


                    foreach (var topic in reportData.MeetingTopics)
                    {
                        foreach (var subtopic in topic.MeetingSubTopics)
                        {
                            foreach (var notes in subtopic.MeetingNotes)
                            {

                                foreach (var files in notes.MeetingNoteFiles)
                                {

                                    if (noteFileCheck == 0)
                                    {
                                        noteFileCheck++;
                                        filestitle.Append(tc000, tc001, tc002, tc003, tc004);
                                        filesTables.Append(filestitle);
                                    }
                                    var tr1 = new TableRow();

                                    // var _date = Convert.ToDateTime(tasks.StartDate).ToString("dd-MM-yyyy") + " to " + Convert.ToDateTime(tasks.EndDate).ToString("dd-MM-yyyy");
                                    var noteNo = topic.TopicNumber + "." + subtopic.SubTopicNumber + "." + notes.NoteNumber;

                                    // Add a cell to each column in the row.
                                    TableCell tc1 = new TableCell(new Paragraph(new Run(new Text(files.Id.ToString()))));
                                    TableCell tc2 = new TableCell(new Paragraph(new Run(new Text(files.Name.Split("~icoPlan~")[1]))));
                                    TableCell tc3 = new TableCell(new Paragraph(new Run(new Text(noteNo))));
                                    TableCell tc4 = new TableCell(new Paragraph(new Run(new Text(files.uploadedByName))));
                                    TableCell tc5 = new TableCell(new Paragraph(new Run(new Text(files.UploadedDate.ToString("dd-MM-yyyy")))));



                                    // Assume you want columns that are automatically sized.
                                    //tc1.Append(new TableCellProperties(
                                    //    new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                                    //tc2.Append(new TableCellProperties(
                                    //   new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2000" }));

                                    //tc3.Append(new TableCellProperties(
                                    //   new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                                    tr1.Append(tc1, tc2, tc3, tc4, tc5);

                                    filesTables.Append(tr1);


                                }
                            }
                        }
                    }

                    mainPart.Document.Body.Append(filesTables);

                }

                GeneratePageNumbers(filepath);

                byte[] returnData = System.IO.File.ReadAllBytes(filepath);


                 
                var data = await _exportReportToNSWordService.ExportReportToMSWord(id, projectId, userId, aDto.JwtId, aDto.Token);

                return File(data, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.ToString());

                return new ContentResult
                {
                    Content = $"Error occured while executing category/{id} -- " + ex.ToString(),
                    ContentType = "text/plain",
                    StatusCode = 500
                };

            }
        }


    }
}
