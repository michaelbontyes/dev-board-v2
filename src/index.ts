import * as dotenv from 'dotenv';
import JiraClient from 'jira-client';
import chalk from 'chalk';

// Load environment variables
dotenv.config();

interface JiraBoard {
  id: number;
  name: string;
  location?: {
    projectKey?: string;
  };
}

interface JiraSprint {
  id: number;
  name: string;
  startDate: string;
  endDate: string;
}

interface JiraWorklog {
  author: {
    displayName: string;
  };
  started: string;
  timeSpentSeconds: number;
}

interface JiraChangelogItem {
  field: string;
  fromString: string;
  toString: string;
  author: {
    displayName: string;
  };
}

interface JiraChangelog {
  created: string;
  items: JiraChangelogItem[];
}

interface IssueCompletionStats {
  [assignee: string]: {
    started: number;
    completed: number;
    completedIssues: string[]; // Array to store completed issue keys
  };
}

interface SprintSummary {
  id: number;
  name: string;
  startDate: string;
  endDate: string;
  totalIssues: number;
  completedIssues: number;
  uatReadyIssues: number;
  timeLogged: {
    [assignee: string]: number;
  };
  completionStats: IssueCompletionStats;
}

interface CommandArgs {
  report: string;
  sprintNumber?: string;
}

// Parse command line arguments
function parseArgs(): CommandArgs {
  const args = process.argv.slice(2);
  const result: CommandArgs = {
    report: 'all' // default report
  };

  for (let i = 0; i < args.length; i++) {
    if (args[i] === 'sprints' && args[i + 1]) {
      result.report = 'sprints';
      result.sprintNumber = args[i + 1];
      break;
    }
  }

  return result;
}

// Initialize Jira client
const jira = new JiraClient({
  protocol: 'https',
  host: process.env.JIRA_HOST?.replace('https://', '') || '',
  username: process.env.JIRA_EMAIL,
  password: process.env.JIRA_API_TOKEN,
  apiVersion: process.env.JIRA_API_VERSION || '2',
  strictSSL: true,
  timeout: Number(process.env.JIRA_REQUEST_TIMEOUT) || 30000
});

async function getSprintIssues(sprintId: number) {
  // Get all issues in the sprint
  const jql = `sprint = ${sprintId}`;
  const allIssues = await jira.searchJira(jql, {
    maxResults: 1000,
    fields: ['summary', 'status', 'assignee', 'timetracking', 'worklog'],
    expand: ['changelog']
  });

  // Get issues that were in UAT Ready during the sprint
  const uatJql = `sprint = ${sprintId} AND status was "UAT Ready"`;
  const uatIssues = await jira.searchJira(uatJql, {
    maxResults: 1000,
    fields: ['key']
  });

  return {
    issues: allIssues.issues,
    total: allIssues.total,
    uatTotal: uatIssues.total
  };
}

async function getAllProjectIssues() {
  try {
    const jql = `project = ${process.env.JIRA_PROJECT_KEY} ORDER BY created DESC`;
    const issues = await jira.searchJira(jql, {
      maxResults: 1000,
      fields: ['summary', 'status', 'assignee', 'priority', 'sprint']
    });
    return issues;
  } catch (error) {
    if (error instanceof Error) {
      console.error('Error fetching issues:', error.message);
    } else {
      console.error('Error fetching issues:', error);
    }
    throw error;
  }
}

function formatDate(dateStr: string): string {
  return new Date(dateStr).toLocaleDateString('en-GB', {
    year: 'numeric',
    month: 'short',
    day: 'numeric'
  });
}

function getCompletionColor(completed: number, total: number): Function {
  const percentage = (completed / total) * 100;
  if (percentage >= 80) return chalk.green;
  if (percentage >= 50) return chalk.yellow;
  return chalk.red;
}

function getTimeLogColor(hours: number): Function {
  if (hours > 60) return chalk.green;
  if (hours > 40) return chalk.yellow;
  return chalk.red;
}

// Add helper function to convert Jira time to hours
function convertJiraTimeToHours(timeSpentSeconds: number): number {
  return timeSpentSeconds / 3600; // Convert seconds to hours
}

async function getAllProjectSprints(sprintNumber?: string) {
  try {
    // First, we need to get the board ID for the project
    const boards = await jira.getAllBoards();
    const projectBoards = boards.values.filter((board: JiraBoard) => 
      board.location?.projectKey === process.env.JIRA_PROJECT_KEY
    );
    
    if (!projectBoards.length) {
      throw new Error(`No board found for project ${process.env.JIRA_PROJECT_KEY}`);
    }

    const board = projectBoards[0];
    console.log('Found board:', board.name, 'with ID:', board.id);

    // Get all sprints for the board
    const sprints = await jira.getAllSprints(board.id);
    
    // Filter sprints if sprint number is provided
    let filteredSprints = sprints.values;
    if (sprintNumber) {
      filteredSprints = sprints.values.filter((sprint: JiraSprint) => 
        sprint.name.includes(sprintNumber)
      );
      
      if (filteredSprints.length === 0) {
        throw new Error(`No sprints found matching number ${sprintNumber}`);
      }
    }
    
    // Get detailed information for each sprint
    const sprintSummaries: SprintSummary[] = [];
    const allAssignees = new Set<string>();
    
    for (const sprint of filteredSprints) {
      console.log(chalk.yellow(`Fetching data for sprint: ${sprint.name}...`));
      const issues = await getSprintIssues(sprint.id);
      const timeLogged: { [key: string]: number } = {};
      const completionStats: IssueCompletionStats = {};
      
      let completedIssues = 0;
      
      for (const issue of issues.issues) {
        if (issue.fields.status.name === 'Done') {
          completedIssues++;
          
          // Process changelog to find who first moved to In Progress
          if (issue.changelog && issue.changelog.histories) {
            let starter = '';
            let firstMoverFromToDo = ''; // Track first person who moved from To Do
            let movedDirectlyToInProgress = false; // Track if moved directly to In Progress
            
            // Sort changelog by date to process events in order
            const sortedHistory = issue.changelog.histories.sort(
              (a: JiraChangelog, b: JiraChangelog) => new Date(a.created).getTime() - new Date(b.created).getTime()
            );

            // First pass: look for direct move to In Progress
            for (const history of sortedHistory) {
              for (const item of history.items) {
                if (item.field === 'status' && item.toString === 'In Progress' && !starter) {
                  starter = history.author.displayName;
                  movedDirectlyToInProgress = item.fromString === 'To Do';
                  break;
                }
                // Track first move from To Do as backup
                if (item.field === 'status' && item.fromString === 'To Do' && !firstMoverFromToDo) {
                  firstMoverFromToDo = history.author.displayName;
                }
              }
              if (starter) break;
            }
            
            // If no one moved it to In Progress, use the first person who moved it from To Do
            if (!starter && firstMoverFromToDo) {
              starter = firstMoverFromToDo;
              movedDirectlyToInProgress = false;
            }
            
            // Record the statistics for whoever we found
            if (starter) {
              completionStats[starter] = completionStats[starter] || { started: 0, completed: 0, completedIssues: [] };
              completionStats[starter].started++;
              completionStats[starter].completed++;
              // Add asterisk if not moved directly to In Progress
              const issueKey = movedDirectlyToInProgress ? issue.key : `${issue.key}*`;
              completionStats[starter].completedIssues.push(issueKey);
            } else {
              // If we still found no one, mark as Unknown
              const unknownStarter = "Unknown";
              completionStats[unknownStarter] = completionStats[unknownStarter] || { started: 0, completed: 0, completedIssues: [] };
              completionStats[unknownStarter].started++;
              completionStats[unknownStarter].completed++;
              completionStats[unknownStarter].completedIssues.push(`${issue.key}*`); // Always add asterisk for unknown
            }
          }
        }
        
        // Reset time logged for this sprint
        if (issue.fields.worklog && issue.fields.worklog.worklogs) {
          // Filter worklogs that fall within the sprint dates
          const sprintStart = new Date(sprint.startDate);
          const sprintEnd = new Date(sprint.endDate);
          
          const sprintWorklogs = issue.fields.worklog.worklogs.filter((worklog: JiraWorklog) => {
            const worklogDate = new Date(worklog.started);
            return worklogDate >= sprintStart && worklogDate <= sprintEnd;
          });

          // Sum up time logged per assignee for this sprint
          for (const worklog of sprintWorklogs) {
            const assignee = worklog.author.displayName;
            allAssignees.add(assignee);
            timeLogged[assignee] = (timeLogged[assignee] || 0) + worklog.timeSpentSeconds;
          }
        }
      }
      
      sprintSummaries.push({
        id: sprint.id,
        name: sprint.name,
        startDate: sprint.startDate,
        endDate: sprint.endDate,
        totalIssues: issues.total,
        completedIssues,
        uatReadyIssues: issues.uatTotal,
        timeLogged,
        completionStats
      });
    }

    // Sort assignees for consistent column order
    const sortedAssignees = Array.from(allAssignees).sort();
    
    // Get max lengths for column sizing
    const maxNameLength = Math.max(...sprintSummaries.map(s => s.name.length), 11);
    const maxDateLength = 12;
    const maxIssuesLength = Math.max(...sprintSummaries.map(s => String(s.totalIssues).length), 5);
    const assigneeColumnWidth = 10;
    
    // Print the table
    console.log('\n' + chalk.bold.blue('Sprint Summary Table:'));
    
    // Calculate total width based on fixed columns and number of assignees
    const headerLine = '─'.repeat(
      maxNameLength + 
      (maxDateLength * 2) + 
      (maxIssuesLength * 3) + // Three columns for Total, Done, and UAT
      8 + // Space for completion percentage
      (assigneeColumnWidth * sortedAssignees.length) + 
      14 // Additional padding
    );
    console.log(chalk.gray(headerLine));
    
    // Print headers
    let header = 
      chalk.bold.white('Sprint'.padEnd(maxNameLength + 2)) +
      chalk.bold.white('Start'.padEnd(maxDateLength + 2)) +
      chalk.bold.white('End'.padEnd(maxDateLength + 2)) +
      chalk.bold.white('Total'.padEnd(maxIssuesLength + 2)) +
      chalk.bold.magenta('UAT'.padEnd(maxIssuesLength + 2)) +
      chalk.bold.white('Done'.padEnd(maxIssuesLength + 2)) +
      chalk.bold.white('%'.padEnd(5));
    
    // Add assignee columns for time logged
    for (const assignee of sortedAssignees) {
      header += chalk.bold.cyan(assignee.substring(0, assigneeColumnWidth).padEnd(assigneeColumnWidth + 2));
    }

    // Print sprint rows
    for (const sprint of sprintSummaries) {
      const completionPercentage = (sprint.completedIssues / sprint.totalIssues) * 100;
      const completionColor = getCompletionColor(sprint.completedIssues, sprint.totalIssues);
      
      let row = 
        chalk.white(sprint.name.padEnd(maxNameLength + 2)) +
        chalk.yellow(formatDate(sprint.startDate).padEnd(maxDateLength + 2)) +
        chalk.yellow(formatDate(sprint.endDate).padEnd(maxDateLength + 2)) +
        chalk.blue(String(sprint.totalIssues).padEnd(maxIssuesLength + 2)) +
        chalk.magenta(String(sprint.uatReadyIssues).padEnd(maxIssuesLength + 2)) +
        completionColor(String(sprint.completedIssues).padEnd(maxIssuesLength + 2)) +
        completionColor(Math.round(completionPercentage) + '%'.padEnd(2));
      
      // Add time logged for each assignee
      for (const assignee of sortedAssignees) {
        const hours = sprint.timeLogged[assignee] 
          ? convertJiraTimeToHours(sprint.timeLogged[assignee])
          : 0;
        const hoursStr = Math.round(hours).toString() + 'h';
        const timeColor = getTimeLogColor(hours);
        row += timeColor(hoursStr.padEnd(assigneeColumnWidth + 2));
      }
      
      console.log(row);
    }
    
    // Add total row
    console.log(chalk.gray(headerLine));
    const totalIssues = sprintSummaries.reduce((sum, s) => sum + s.totalIssues, 0);
    const totalCompleted = sprintSummaries.reduce((sum, s) => sum + s.completedIssues, 0);
    const totalUatReady = sprintSummaries.reduce((sum, s) => sum + s.uatReadyIssues, 0);
    const totalCompletionPercentage = (totalCompleted / totalIssues) * 100;
    const totalCompletionColor = getCompletionColor(totalCompleted, totalIssues);

    let totalRow = 
      chalk.bold.white('TOTAL'.padEnd(maxNameLength + 2)) +
      ''.padEnd(maxDateLength + 2) +
      ''.padEnd(maxDateLength + 2) +
      chalk.bold.blue(String(totalIssues).padEnd(maxIssuesLength + 2)) +
      chalk.bold.magenta(String(totalUatReady).padEnd(maxIssuesLength + 2)) +
      totalCompletionColor(String(totalCompleted).padEnd(maxIssuesLength + 2)) +
      totalCompletionColor(Math.round(totalCompletionPercentage) + '%'.padEnd(2));

    // Add total hours per assignee
    for (const assignee of sortedAssignees) {
      const totalHours = convertJiraTimeToHours(sprintSummaries.reduce((sum, sprint) => 
        sum + (sprint.timeLogged[assignee] || 0), 0));
      const hoursStr = Math.round(totalHours).toString() + 'h';
      const timeColor = getTimeLogColor(totalHours);
      totalRow += timeColor(hoursStr.padEnd(assigneeColumnWidth + 2));
    }
    console.log(totalRow);
    console.log(chalk.gray(headerLine));

    // After printing the main table, add the Issue Completion table
    console.log('\n' + chalk.bold.blue('Issue Completion Table:'));
    
    // Calculate total width for the new table
    const completionHeaderLine = '─'.repeat(maxNameLength + 40);
    console.log(chalk.gray(completionHeaderLine));
    
    // Print completion table headers
    const completionHeader = 
      chalk.bold.white('Sprint'.padEnd(maxNameLength + 2)) +
      chalk.bold.white('Assignee'.padEnd(20)) +
      chalk.bold.white('Started'.padEnd(10)) +
      chalk.bold.white('Done'.padEnd(8)) +
      chalk.bold.white('Completed Issues');
    
    console.log(completionHeader);
    console.log(chalk.gray(completionHeaderLine));
    
    // Print completion stats for each sprint
    for (const sprint of sprintSummaries) {
      const assignees = Object.keys(sprint.completionStats).sort();
      
      for (const assignee of assignees) {
        const stats = sprint.completionStats[assignee];
        const row = 
          chalk.white(sprint.name.padEnd(maxNameLength + 2)) +
          chalk.cyan(assignee.padEnd(20)) +
          chalk.yellow(String(stats.started).padEnd(10)) +
          getCompletionColor(stats.completed, stats.started)(String(stats.completed).padEnd(8)) +
          chalk.gray(stats.completedIssues.join(', '));
        
        console.log(row);
      }
      
      // Add a separator line between sprints
      console.log(chalk.gray(completionHeaderLine));
    }
    
    // Add total row for completion stats
    const totalStats: IssueCompletionStats = {};
    for (const sprint of sprintSummaries) {
      for (const [assignee, stats] of Object.entries(sprint.completionStats)) {
        totalStats[assignee] = totalStats[assignee] || { started: 0, completed: 0, completedIssues: [] };
        totalStats[assignee].started += stats.started;
        totalStats[assignee].completed += stats.completed;
        totalStats[assignee].completedIssues = totalStats[assignee].completedIssues.concat(stats.completedIssues);
      }
    }
    
    // Print total stats with completed issues
    const sortedTotalAssignees = Object.keys(totalStats).sort();
    for (const assignee of sortedTotalAssignees) {
      const stats = totalStats[assignee];
      const row = 
        chalk.bold.white('TOTAL'.padEnd(maxNameLength + 2)) +
        chalk.cyan(assignee.padEnd(20)) +
        chalk.yellow(String(stats.started).padEnd(10)) +
        getCompletionColor(stats.completed, stats.started)(String(stats.completed).padEnd(8)) +
        chalk.gray(stats.completedIssues.join(', '));
      
      console.log(row);
    }
    console.log(chalk.gray(completionHeaderLine));

    return sprintSummaries;
  } catch (error) {
    if (error instanceof Error) {
      console.error(chalk.red('Error fetching sprints:', error.message));
    } else {
      console.error(chalk.red('Error fetching sprints:', error));
    }
    throw error;
  }
}

async function main() {
  try {
    const args = parseArgs();
    
    switch (args.report) {
      case 'sprints':
        console.log(chalk.blue(`\nFetching sprint data${args.sprintNumber ? ` for sprint ${args.sprintNumber}` : ''}...`));
        await getAllProjectSprints(args.sprintNumber);
        break;
      default:
        console.log(chalk.blue('\nFetching all project data...'));
        await getAllProjectSprints();
    }
  } catch (error) {
    if (error instanceof Error) {
      console.error(chalk.red('Error in main execution:', error.message));
    } else {
      console.error(chalk.red('Error in main execution:', error));
    }
    process.exit(1);
  }
}

// Add usage information
if (process.argv.includes('--help') || process.argv.includes('-h')) {
  console.log(`
${chalk.bold('Usage:')}
  npm start                    - Show all sprints
  npm start sprints <number>   - Show specific sprint(s) containing <number>
  npm start --help             - Show this help message

${chalk.bold('Examples:')}
  npm start sprints 21         - Show sprint containing "21"
  npm start                    - Show all sprints
  `);
  process.exit(0);
}

main();
