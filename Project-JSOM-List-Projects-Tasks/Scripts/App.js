/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
'use strict';

SP.SOD.executeOrDelayUntilScriptLoaded(initializePage, "PS.js");

//Project PWA Context and published projects in PWA
var projContext;
var projects;

function initializePage() {

    //Get the Project context for this web
    projContext = PS.ProjectContext.get_current();

    loadProjects();
}

//General CSOM failure event handler
//Invoked when ExecuteQueryAsync returns unsuccessfully.
function onRequestFailed(sender, args) {
    alert("Failed to execute: " + args.get_message());
    return;
};

//Query CSOM and get the list of projects in PWA
function loadProjects() {

    projects = projContext.get_projects();

    //Request to server - identifies what to retrieve
    projContext.load(projects, 'Include(Name, Id)');

    //Notice to server to execute query
    projContext.executeQueryAsync(displayProjects, onRequestFailed);

    // Syntax for requesting more fields to pull down from server
    // projContext.load(projects, 'Include(Name, Description, StartDate, Id, IsCheckedOut)');
}

//Display the projects with names and ids in a table */
function displayProjects() {

    //Current published project and ID
    var p, projId;

    //Project table rows to publish collectively
    var pTable = [];

    var pEnum = projects.getEnumerator();

    //Build a 3-column table with one project per row.
    while (pEnum.moveNext()) {
        p = pEnum.get_current();

        //Build row and add it to the end of the table.
        pTable = buildProjectRow(pTable, p);
    }

    //Append table as the HTML table body
    $("tbody").append(pTable);

    document.getElementById("msg").innerHTML = "Enumerated projects = " + pTable.length;
}


/* Query CSOM and get the list of tasks for a specific project  */
function btnLoadTasks(pid) {
    //Event handler for the "Show tasks" buttons. 
    //The project ID is the sole argument and is used to get the appropriate task info from the service.
    //The project ID is also the button name, and is used to identify where to place the task information in the table.
    
    //Project ID to pass to the event handler
    var projId = pid;

    //Get the project reference
    var pProj = projects.getById(projId);

    //Get the tasks collection reference associated with the project.
    var tasks = pProj.get_tasks();

    projContext.load(tasks, 'Include(Id, Name, Start, ScheduledStart, Completion)');

    // If successful, let displayTasks handle the presentation details
    projContext.executeQueryAsync(function () { displayTasks(tasks, projId) }, onRequestFailed);
}

/* Insert tasks for the specified project immediately under the project entry in the table.*/
function displayTasks(tasks, projId) {

    //selected project ID
    var pId = projId;

    //individual task
    var t;                               

    //Task table rows to publish collectively
    var tTable = [];                                

    var tEnum = tasks.getEnumerator();

    //Build table one task per row.
    while (tEnum.moveNext()) {
        t = tEnum.get_current();

        //Build row and add it to the end of the table.
        tTable = buildTaskRow(tTable, t);
    }

    //Locate row in the HTML table using the button name. Insert tasks after this row. 
    var button = $("button[name='" + pId + "']");
    button.parents("tr").after(tTable);

    document.getElementById("msg").innerHTML = "Task count: " + tTable.length;
}

/* Zap all tasks from the Project/Task table. */
function btnRemoveTasks() {
    $(".tasks").Remove();
}

/* Build HTML table row for Project entry           */
function buildProjectRow(pTable, p) {

    //Builds a table of 3 cells.
    //    * cell/column 1 contains a button. the button caals the event handler that lists tasks.
    //      The button name is the project ID, and is used to locate the table row immediately 
    //      above the insertion.
    //    * cell/column 2 contains the project name.
    //    * cell/column 3 contains the project id.

    //Current published project object, and ID and name
    var project = p;
    var projId = p.get_id();
    var projName = p.get_name();

    //Output table row
    var tblRow = ""; 

    const TRTAG = "<tr  class='proj'>";
    const TREND = "</tr>";

    const CELLTAG = "<td>";
    const WIDECELLTAG = "<td width='300px'>";
    const CELLEND = "</td>";

    const BTNTAG = "<button height='20px' name='";    // Set name to p.get_id() 
    const BTNMID = "' onclick='btnLoadTasks(\"";      // Pass fn argument p.get_id()
    const BTNEND = "\"); return false;'>Show tasks</button>";

    // table row for individual project
    tblRow = TRTAG;

    //Column 1 - Set button name to pprojId Name, and use projId for the function argument value.
    tblRow += CELLTAG + BTNTAG + projId + BTNMID + projId + BTNEND + CELLEND;

    //Column 2 - Set displayed value to pprojName
    tblRow += WIDECELLTAG + projName + CELLEND;

    //Column 3 - Set displayed value to pprojId
    tblRow += CELLTAG + projId + CELLEND;

    tblRow += TREND;

    pTable.push(tblRow);

    return pTable;
}

/* Build HTML table row for task entry                         */
function buildTaskRow(tTable, t) {

    //This function builds an array of tasks associated with a project. 
    //The table row has 3 cells/columns. 
    //    - The first cell is blank to distinguish it from a row containing project info.
    //    - The second cell contains the task name.
    //    - The third cell contains the task ID.

    var task = t;
    var taskId = t.get_id();
    var taskName = t.get_name();

    //Row to append to the table
    var tTblRow = "";

    const TRTAGT = "<tr  class='task'>";
    const TREND = "</tr>";

    const CELLTAG = "<td>";
    const WIDECELLTAG = "<td width='240px'>";
    const CELLEND = "</td>";

    tTblRow = TRTAGT;

    // Columns 1, 2, and 3. One per line.
    tTblRow += CELLTAG + CELLEND;
    tTblRow += WIDECELLTAG + taskName + CELLEND;
    tTblRow += CELLTAG + taskId + CELLEND;

    tTblRow += TREND;

    tTable.push(tTblRow);
        
    return tTable;
}


