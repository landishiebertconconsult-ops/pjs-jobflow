import { useMemo, useState, type CSSProperties, type ReactNode } from "react";
import * as XLSX from "xlsx";
import { BarChart3, BriefcaseBusiness, CalendarClock, CheckSquare, ClipboardCheck, ClipboardList, FileText, LayoutDashboard, LogOut, PackageCheck, Plus, Search, Settings, ShieldCheck, Timer, Trash2, Users } from "lucide-react";

type Role = "Admin" | "Project Manager" | "Foreman" | "Crew";
type View =
  | "dashboard"
  | "planning"
  | "reports"
  | "crewTasks"
  | "crewPoints"
  | "inspections"
  | "timeTracking"
  | "labourTracking"
  | "potentialJobs"
  | "calendar"
  | "smallJobs";
type JobFolder = "crewTasks" | "inspections" | "pmTasks" | "dailyNotes" | "photos" | "documents" | "knowifyUploads" | "activityLog";
type JobDocument = { id: number; jobId: number; fileName: string; uploadedBy: string; uploadedDate: string; uploadedAt?: string; type: "Photo" | "Document"; url?: string };
type JobDocumentUpload = { fileName: string; type: "Photo" | "Document"; url?: string };
type JobStatus = "Planning" | "Active" | "On Hold" | "Complete";
type Risk = "Green" | "Yellow" | "Red";
type Task = { id: number; title: string; done: boolean };
type Inspection = { id: number; title: string; status: string; date: string; notes: string };

type User = {
  id: number;
  name: string;
  role: Role;
  points: number;
  monthlyPointLimit: number;
  email?: string;
  password?: string;
  mustSetPassword?: boolean;
  active?: boolean;
  lastLogin?: string;
};

type Job = {
  id: number;
  jobNumber: string;
  name: string;
  customer: string;
  location: string;
  status: JobStatus;
  certainty: string;
  phase: string;
  risk: Risk;
  startDate: string;
  finishDate: string;
  nextAction: string;
  actionOwner: string;
  actionDue: string;
  progress: number;
  allowedHours: number;
  usedHours: number;
  labourBudget: number;
  labourCostToDate: number;
  budget: number;
  costToDate: number;
  crewIds: number[];
  crewTasks: Task[];
  projectTasks: Task[];
  pmRequests: Task[];
  inspections: Inspection[];
};

type SmallJob = {
  id: number;
  title: string;
  location: string;
  estimatedDays: number;
  estimatedCrew: number;
  priority: "Low" | "Medium" | "High" | "Urgent";
  scope: string;
  documents: JobDocumentUpload[];
  status: "Not Started" | "Planning" | "In Progress" | "Completed";
  enteredDate: string;
  scheduled: boolean;
};

type PotentialJobActivity = {
  id: number;
  date: string;
  user: string;
  field: string;
  previousValue: string;
  newValue: string;
};

type PotentialJob = {
  id: number;
  name: string;
  customer: string;
  location: string;
  probability: string;
  startDate: string;
  finishDate: string;
  crewNeeded: number;
  estimatedHours: number;
  assumedValue: number;
  crewIds: number[];
  document: string;
  documents: JobDocumentUpload[];
  activityLog: PotentialJobActivity[];
};

type PointHistory = {
  id: number;
  userId: number;
  awardedById: number;
  points: number;
  reason: string;
  date: string;
  jobName: string;
};

type DailyReport = {
  id: number;
  jobId: number;
  userId: number;
  userName: string;
  jobName: string;
  note: string;
  date: string;
  pointsAwarded: number;
  pictures: string[];
  pictureUploads?: JobDocumentUpload[];
};

const usersSeed: User[] = [
  { id: 1, name: "Landis Hiebert", role: "Admin", points: 0, monthlyPointLimit: 999, email: "landis@pjselectric.ca", password: "admin", mustSetPassword: false, active: true },
  { id: 2, name: "Jason Giesbrecht", role: "Project Manager", points: 0, monthlyPointLimit: 999, email: "jason@pjselectric.ca", password: "temp123", mustSetPassword: true },
  { id: 3, name: "Riley Pakosh", role: "Foreman", points: 35, monthlyPointLimit: 25, email: "riley@pjselectric.ca", password: "temp123", mustSetPassword: true },
  { id: 4, name: "Miguel Kehler", role: "Foreman", points: 20, monthlyPointLimit: 25, email: "miguel@pjselectric.ca", password: "temp123", mustSetPassword: true },
  { id: 5, name: "Kris Sawyer", role: "Crew", points: 15, monthlyPointLimit: 0, email: "kris@pjselectric.ca", password: "temp123", mustSetPassword: true },
  { id: 6, name: "Jake Turner", role: "Crew", points: 5, monthlyPointLimit: 0, email: "jake@pjselectric.ca", password: "temp123", mustSetPassword: true },
  { id: 7, name: "Ethan Funk", role: "Crew", points: 0, monthlyPointLimit: 0, email: "ethan@pjselectric.ca", password: "temp123", mustSetPassword: true },
];

const pmTaskList = (completeThrough: number): Task[] => [
  "Project awarded / contract received",
  "Permit pulled",
  "Kickoff meeting complete",
  "Shop drawings submitted",
  "Shop drawings approved",
  "Long-lead material ordered",
  "Site mobilization complete",
  "Service rough-in complete",
  "Lighting rough-in complete",
  "Distribution rough-in complete",
  "Fire alarm rough-in complete",
  "Low-voltage rough-in complete",
  "Ceiling close-in complete",
  "Device installation complete",
  "Lighting installation complete",
  "Panel schedules / labels complete",
  "Testing and commissioning complete",
  "Deficiency walkthrough complete",
  "As-builts / O&M manuals submitted",
].map((title, index) => ({ id: index + 1, title, done: index < completeThrough }));

const initialJobs: Job[] = [
  {
    id: 1,
    jobNumber: "PJ-2026-001",
    name: "Fort Hope Service / Temp Power",
    customer: "Penn-co Construction",
    location: "Fort Hope",
    status: "Active",
    certainty: "Confirmed",
    phase: "Service Installation",
    risk: "Yellow",
    startDate: "2026-05-01",
    finishDate: "2026-06-15",
    nextAction: "Splice additional 500 MCM cable and bring to volt shack",
    actionOwner: "Riley Pakosh",
    actionDue: "2026-05-20",
    progress: 55,
    allowedHours: 1200,
    usedHours: 670,
    labourBudget: 92000,
    labourCostToDate: 51400,
    budget: 185000,
    costToDate: 91250,
    crewIds: [3, 4, 5],
    crewTasks: [
      { id: 1, title: "Submit daily notes", done: false },
      { id: 2, title: "Upload progress photos", done: true },
      { id: 3, title: "Update material list", done: false },
    ],
    projectTasks: pmTaskList(7),
    pmRequests: [
      { id: 1, title: "Confirm additional 500 MCM splice material — requested by Riley Pakosh", done: false },
      { id: 2, title: "Send updated next-action direction to site — requested by Kris Sawyer", done: false },
    ],
    inspections: [
      { id: 1, title: "Service inspection", status: "Pending", date: "2026-05-22", notes: "Waiting for Hydro confirmation" },
      { id: 2, title: "Ground inspection", status: "Scheduled", date: "2026-05-21", notes: "Ground plate photo required" },
      { id: 3, title: "Final inspection", status: "Not Started", date: "", notes: "After camp tie-in" },
    ],
  },
  {
    id: 2,
    jobNumber: "PJ-2026-002",
    name: "Aroland Elders Lodge",
    customer: "GC / Owner",
    location: "Aroland",
    status: "Active",
    certainty: "Confirmed",
    phase: "Commissioning",
    risk: "Green",
    startDate: "2026-01-10",
    finishDate: "2026-06-30",
    nextAction: "Complete commissioning and closeout documents",
    actionOwner: "Miguel Kehler",
    actionDue: "2026-05-24",
    progress: 82,
    allowedHours: 2400,
    usedHours: 1968,
    labourBudget: 168000,
    labourCostToDate: 137760,
    budget: 350000,
    costToDate: 282000,
    crewIds: [4, 5],
    crewTasks: [
      { id: 1, title: "Panel photos complete", done: true },
      { id: 2, title: "Submit deficiency notes", done: false },
    ],
    projectTasks: pmTaskList(16),
    pmRequests: [{ id: 1, title: "Review commissioning documents — requested by Miguel Kehler", done: false }],
    inspections: [
      { id: 1, title: "Rough-in inspection", status: "Passed", date: "2026-04-02", notes: "No open items" },
      { id: 2, title: "Final inspection", status: "Pending", date: "2026-05-28", notes: "Book after emergency lights are tested" },
    ],
  },
];

const initialPotentialJobs: PotentialJob[] = [
  {
    id: 1,
    name: "Potential Summer School Retrofit",
    customer: "TBD",
    location: "Manitoba",
    probability: "Medium",
    startDate: "2026-06-01",
    finishDate: "2026-07-15",
    crewNeeded: 4,
    estimatedHours: 1200,
    assumedValue: 175000,
    crewIds: [],
    document: "Preliminary drawings.pdf",
    documents: [{ fileName: "Preliminary drawings.pdf", type: "Document" }],
    activityLog: [
      {
        id: 1,
        date: "2026-05-22",
        user: "System",
        field: "Original Entry",
        previousValue: "—",
        newValue: "Initial potential job created",
      },
    ],
  },
  {
    id: 2,
    name: "Northern Camp Expansion",
    customer: "Penn-co Construction",
    location: "Northern Manitoba",
    probability: "High",
    startDate: "2026-07-01",
    finishDate: "2026-10-01",
    crewNeeded: 6,
    estimatedHours: 2200,
    assumedValue: 325000,
    crewIds: [],
    document: "Tender set.pdf",
    documents: [{ fileName: "Tender set.pdf", type: "Document" }],
    activityLog: [
      {
        id: 2,
        date: "2026-05-22",
        user: "System",
        field: "Original Entry",
        previousValue: "—",
        newValue: "Initial potential job created",
      },
    ],
  },
];

const initialPoints: PointHistory[] = [
  { id: 1, userId: 3, awardedById: 1, points: 25, reason: "Stayed late to finish service prep", date: "2026-05-15", jobName: "Fort Hope Service / Temp Power" },
  { id: 2, userId: 5, awardedById: 3, points: 15, reason: "Kept material organized", date: "2026-05-13", jobName: "Fort Hope Service / Temp Power" },
  { id: 3, userId: 4, awardedById: 1, points: 20, reason: "Resolved inspection item quickly", date: "2026-05-14", jobName: "Aroland Elders Lodge" },
];

const initialDailyReports: DailyReport[] = [
  { id: 1, jobId: 1, userId: 5, userName: "Kris Sawyer", jobName: "Fort Hope Service / Temp Power", note: "Submitted daily notes and progress photos.", date: "2026-05-13", pointsAwarded: 5, pictures: ["progress-photo-1.jpg"] },
];

const initialDocuments: JobDocument[] = [
  { id: 1, jobId: 1, fileName: "volt-shack-progress.jpg", uploadedBy: "Kris Sawyer", uploadedDate: "2026-05-13", type: "Photo" },
  { id: 2, jobId: 1, fileName: "ground-plate-photo.jpg", uploadedBy: "Riley Pakosh", uploadedDate: "2026-05-14", type: "Photo" },
];

const pct = (used: number, total: number) => Math.max(0, Math.min(100, Math.round((used / Math.max(total, 1)) * 100)));
const initials = (name: string) => name.split(" ").map((n) => n[0]).join("").slice(0, 2);
const userById = (users: User[], id: number) => users.find((u) => u.id === id);
const currency = (n: number) => `$${Math.round(n).toLocaleString()}`;
const budgetUsedPctForJob = (job: Job) => pct(job.costToDate, job.budget);
const budgetLeftPctForJob = (job: Job) => Math.max(0, 100 - budgetUsedPctForJob(job));
const budgetPriorityRank = (job: Job) => {
  const used = budgetUsedPctForJob(job);
  const left = budgetLeftPctForJob(job);
  if (used >= 100) return 0;
  if (left <= 10) return 1;
  if (left <= 25) return 2;
  return 3;
};
const budgetPriorityLabel = (job: Job) => {
  const used = budgetUsedPctForJob(job);
  const left = budgetLeftPctForJob(job);
  if (used >= 100) return "⚠ At / Over Budget";
  if (left <= 10) return "⚠ Within 10%";
  if (left <= 25) return "Caution: Within 25%";
  return "OK";
};
const budgetPriorityTone = (job: Job): "green" | "yellow" | "red" => {
  const rank = budgetPriorityRank(job);
  if (rank === 0 || rank === 1) return "red";
  if (rank === 2) return "yellow";
  return "green";
};

export default function App() {
  const [activeView, setActiveView] = useState<View>("dashboard");
  const [users, setUsers] = useState<User[]>(usersSeed);
  const [jobs, setJobs] = useState<Job[]>(initialJobs);
  const [potentialJobs, setPotentialJobs] = useState<PotentialJob[]>(initialPotentialJobs);
  const [smallJobs, setSmallJobs] = useState<SmallJob[]>([
    {
      id: 1,
      title: "Replace parking lot pole light",
      location: "Steinbach Shop",
      estimatedDays: 1,
      estimatedCrew: 2,
      priority: "Medium",
      scope: "Replace damaged pole light fixture and reconnect controls.",
      documents: [],
      status: "Planning",
      enteredDate: new Date().toISOString().slice(0, 10),
      scheduled: false,
    },
  ]);
  const [points, setPoints] = useState<PointHistory[]>(initialPoints);
  const [dailyReports, setDailyReports] = useState<DailyReport[]>(initialDailyReports);
  const [documents, setDocuments] = useState<JobDocument[]>(initialDocuments);
  const [selectedId, setSelectedId] = useState(1);
  const [search, setSearch] = useState("");
  const [currentUserId, setCurrentUserId] = useState(1);
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [loginEmail, setLoginEmail] = useState("landis@pjselectric.ca");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [passwordSetupUserId, setPasswordSetupUserId] = useState<number | null>(null);
  const [showSettings, setShowSettings] = useState(false);
  const [showAddJob, setShowAddJob] = useState(false);

  const currentUser = userById(users, currentUserId) || users[0];
  const isAdmin = currentUser.role === "Admin" || currentUser.role === "Project Manager";
  const isForeman = currentUser.role === "Foreman";

  const visibleJobs = useMemo(() => {
    const allowed = isAdmin ? jobs : jobs.filter((job) => job.crewIds.includes(currentUser.id));
    return allowed
      .filter((job) => `${job.jobNumber} ${job.name} ${job.customer} ${job.location}`.toLowerCase().includes(search.toLowerCase()))
      .sort((a, b) => budgetPriorityRank(a) - budgetPriorityRank(b) || budgetLeftPctForJob(a) - budgetLeftPctForJob(b));
  }, [jobs, isAdmin, currentUser.id, search]);

  const selectedJob = visibleJobs.find((job) => job.id === selectedId) || visibleJobs[0] || jobs[0];
  const selectedCrew = users.filter((u) => selectedJob.crewIds.includes(u.id));

  function updateJob(update: Partial<Job>) {
    setJobs((all) => all.map((job) => (job.id === selectedJob.id ? { ...job, ...update } : job)));
  }

  function toggleTask(type: "crewTasks" | "projectTasks" | "pmRequests", id: number) {
    setJobs((all) =>
      all.map((job) =>
        job.id === selectedJob.id
          ? { ...job, [type]: job[type].map((task) => (task.id === id ? { ...task, done: !task.done } : task)) }
          : job
      )
    );
  }

  function updateInspection(inspectionId: number, status: string) {
    setJobs((all) =>
      all.map((job) =>
        job.id === selectedJob.id
          ? { ...job, inspections: job.inspections.map((inspection) => (inspection.id === inspectionId ? { ...inspection, status } : inspection)) }
          : job
      )
    );
  }

  function awardPoints(userId: number, amount: number, reason: string, jobNameOverride?: string) {
    if (!userId || amount <= 0 || !reason.trim()) return;
    if (isForeman) {
      const awarded = points.filter((p) => p.awardedById === currentUser.id).reduce((sum, p) => sum + p.points, 0);
      if (awarded + amount > currentUser.monthlyPointLimit) return;
    }
    setUsers((all) => all.map((u) => (u.id === userId ? { ...u, points: u.points + amount } : u)));
    setPoints((all) => [{ id: Date.now(), userId, awardedById: currentUser.id, points: amount, reason, date: new Date().toISOString().slice(0, 10), jobName: jobNameOverride || selectedJob.name }, ...all]);
  }

  function deductPoints(userId: number, amount: number, reason: string) {
    if (!isAdmin || !userId || amount <= 0 || !reason.trim()) return;
    const targetUser = users.find((user) => user.id === userId);
    if (!targetUser) return;
    const deduction = Math.min(amount, targetUser.points);
    setUsers((all) => all.map((user) => user.id === userId ? { ...user, points: Math.max(0, user.points - deduction) } : user));
    setPoints((all) => [{ id: Date.now(), userId, awardedById: currentUser.id, points: -deduction, reason: `Redeemed / Purchase: ${reason}`, date: new Date().toISOString().slice(0, 10), jobName: "Crew Points Store" }, ...all]);
  }

  function toggleCrewTaskForJob(jobId: number, taskId: number) {
    setJobs((all) => all.map((job) => job.id === jobId ? { ...job, crewTasks: job.crewTasks.map((task) => task.id === taskId ? { ...task, done: !task.done } : task) } : job));
  }

  function addCrewTaskForJob(jobId: number, title: string) {
    if (!title.trim()) return;
    setJobs((all) => all.map((job) => job.id === jobId ? { ...job, crewTasks: [{ id: Date.now(), title: `${title.trim()} — added by ${currentUser.name}`, done: false }, ...job.crewTasks] } : job));
  }

  function submitDailyReport(jobId: number, note: string, pictures: string[] = [], pictureUploads: JobDocumentUpload[] = []) {
    if (!note.trim()) return;
    const job = jobs.find((item) => item.id === jobId);
    if (!job) return;

    const today = new Date().toISOString().slice(0, 10);
    const canEarnDailyPoints = currentUser.role === "Crew" || currentUser.role === "Foreman";
    const existingTodayReport = dailyReports.find((report) => report.jobId === jobId && report.userId === currentUser.id && report.date === today);
    const alreadyEarnedDailyReportPoints = points.some((point) => point.userId === currentUser.id && point.date === today && point.reason === "Daily report submitted");
    const alreadyEarnedPhotoBonus = points.some((point) => point.userId === currentUser.id && point.date === today && point.reason === "Daily report photo bonus");
    const dailyPointsAwarded = canEarnDailyPoints && !alreadyEarnedDailyReportPoints ? 5 : 0;
    const photoBonusAwarded = canEarnDailyPoints && !alreadyEarnedPhotoBonus && pictureUploads.length >= 5 ? 5 : 0;
    const pointsAwarded = dailyPointsAwarded + photoBonusAwarded;

    const finalPictures = pictures.length ? pictures : existingTodayReport?.pictures || [];
    const finalPictureUploads = pictureUploads.length ? pictureUploads : existingTodayReport?.pictureUploads || [];

    const reportPayload: DailyReport = {
      id: existingTodayReport?.id || Date.now(),
      jobId,
      userId: currentUser.id,
      userName: currentUser.name,
      jobName: job.name,
      note: note.trim(),
      date: today,
      pointsAwarded: existingTodayReport ? existingTodayReport.pointsAwarded + photoBonusAwarded : pointsAwarded,
      pictures: finalPictures,
      pictureUploads: finalPictureUploads,
    };

    setDailyReports((all) => {
      if (existingTodayReport) {
        return all.map((report) => report.id === existingTodayReport.id ? reportPayload : report);
      }
      return [reportPayload, ...all];
    });

    const uploadsForJobFile = finalPictureUploads.length
      ? finalPictureUploads
      : finalPictures.map((fileName) => ({ fileName, type: "Photo" as const }));

    if (uploadsForJobFile.length > 0) {
      setDocuments((all) => {
        const newUploads = uploadsForJobFile.filter((upload) => !all.some((document) => document.jobId === jobId && document.uploadedDate === today && document.uploadedBy === currentUser.name && document.fileName === upload.fileName));

        if (newUploads.length === 0) return all;

        return [
          ...newUploads.map((upload, index) => ({
            id: Date.now() + index + 10,
            jobId,
            fileName: upload.fileName,
            uploadedBy: currentUser.name,
            uploadedDate: today,
            uploadedAt: new Date().toLocaleString(),
            type: upload.type,
            url: upload.url,
          })),
          ...all,
        ];
      });
    }

    if (pointsAwarded > 0) {
      setUsers((all) => all.map((user) => user.id === currentUser.id ? { ...user, points: user.points + pointsAwarded } : user));
      const newPointRows: PointHistory[] = [];
      if (dailyPointsAwarded > 0) {
        newPointRows.push({ id: Date.now() + 1, userId: currentUser.id, awardedById: currentUser.id, points: dailyPointsAwarded, reason: "Daily report submitted", date: today, jobName: job.name });
      }
      if (photoBonusAwarded > 0) {
        newPointRows.push({ id: Date.now() + 2, userId: currentUser.id, awardedById: currentUser.id, points: photoBonusAwarded, reason: "Daily report photo bonus", date: today, jobName: job.name });
      }
      setPoints((all) => [...newPointRows, ...all]);
    }
  }

  function uploadJobDocuments(jobId: number, uploads: JobDocumentUpload[]) {
    if (uploads.length === 0) return;
    const today = new Date().toISOString().slice(0, 10);
    setDocuments((all) => [
      ...uploads.map((upload, index) => ({
        id: Date.now() + index,
        jobId,
        fileName: upload.fileName,
        uploadedBy: currentUser.name,
        uploadedDate: today,
        uploadedAt: new Date().toLocaleString(),
        type: upload.type,
        url: upload.url,
      })),
      ...all,
    ]);
  }

  function addJob() {
    setShowAddJob(true);
  }

  function deleteJob(jobId: number) {
    setJobs((prevJobs) => {
      const remainingJobs = prevJobs.filter((job) => job.id !== jobId);
      setSelectedId(remainingJobs[0]?.id || 0);
      return remainingJobs;
    });
  }

  if (!isLoggedIn) {
    const setupUser = passwordSetupUserId ? users.find((user) => user.id === passwordSetupUserId) : null;
    return (
      <>
        <MobileCrewStyles />
        {setupUser ? (
          <PasswordSetupPage
            user={setupUser}
            onSave={(newPassword) => {
              setUsers((all) => all.map((user) => user.id === setupUser.id ? { ...user, password: newPassword, mustSetPassword: false } : user));
              setCurrentUserId(setupUser.id);
              setIsLoggedIn(true);
              setPasswordSetupUserId(null);
              setLoginPassword("");
              setLoginError("");
            }}
            onCancel={() => setPasswordSetupUserId(null)}
          />
        ) : (
          <LoginPage
            loginEmail={loginEmail}
            setLoginEmail={setLoginEmail}
            loginPassword={loginPassword}
            setLoginPassword={setLoginPassword}
            loginError={loginError}
            onLogin={() => {
              const normalizedEmail = loginEmail.trim().toLowerCase();
              const loginUser = users.find((user) => (user.email || "").toLowerCase() === normalizedEmail);
              if (!loginUser) {
                setLoginError("No account found for that email address.");
                return;
              }
              if (loginUser.active === false) {
                setLoginError("This account is inactive. Contact Admin/PM to reactivate it.");
                return;
              }
              if ((loginUser.password || "") !== loginPassword) {
                if (!(loginUser.role === "Admin" && loginPassword === "admin")) {
                  setLoginError("Incorrect password. For the prototype, Admin password is admin. Crew/Foreman temporary password is temp123.");
                  return;
                }
              }
              if (loginUser.mustSetPassword) {
                setPasswordSetupUserId(loginUser.id);
                setLoginError("");
                return;
              }
              setCurrentUserId(loginUser.id);
              setIsLoggedIn(true);
              setLoginPassword("");
              setLoginError("");
            }}
          />
        )}
      </>
    );
  }

  return (
    <div id="jobflow-app" style={styles.app}>
      <MobileCrewStyles />
      <Sidebar activeView={activeView} setActiveView={setActiveView} user={currentUser} isAdmin={isAdmin} onAddJob={addJob} onOpenSettings={() => setShowSettings(true)} onSignOut={() => setIsLoggedIn(false)} />
      <main style={styles.main}>
        <Header currentUser={currentUser} onOpenSettings={() => setShowSettings(true)} />

        {activeView === "dashboard" && (
          <DashboardView jobs={visibleJobs} selectedJob={selectedJob} selectedCrew={selectedCrew} search={search} setSearch={setSearch} setSelectedId={setSelectedId} updateJob={updateJob} deleteJob={deleteJob} toggleTask={toggleTask} updateInspection={updateInspection} isAdmin={isAdmin} currentUser={currentUser} dailyReports={dailyReports}
            submitDailyReport={submitDailyReport}
            documents={documents}
            uploadJobDocuments={uploadJobDocuments}
            potentialJobs={potentialJobs}
            setPotentialJobs={setPotentialJobs} />
        )}
        {activeView === "planning" && <PlanningView isAdmin={isAdmin} jobs={jobs} visibleJobs={visibleJobs} potentialJobs={potentialJobs} setPotentialJobs={setPotentialJobs} users={users} />}
        {activeView === "reports" && <ReportsView isAdmin={isAdmin} jobs={visibleJobs} />}
        {activeView === "crewTasks" && <CrewTasksView jobs={visibleJobs} currentUser={currentUser} toggleCrewTaskForJob={toggleCrewTaskForJob} addCrewTaskForJob={addCrewTaskForJob} submitDailyReport={submitDailyReport} dailyReports={dailyReports} />}
        {activeView === "crewPoints" && <CrewPoints users={users} setUsers={setUsers} jobs={jobs} points={points} currentUser={currentUser} selectedJob={selectedJob} awardPoints={awardPoints} deductPoints={deductPoints} isAdmin={isAdmin} isForeman={isForeman} />}
        {activeView === "potentialJobs" && <PotentialJobsPage isAdmin={isAdmin} currentUser={currentUser} potentialJobs={potentialJobs} setPotentialJobs={setPotentialJobs} />}
        {activeView === "calendar" && <CalendarView jobs={visibleJobs} potentialJobs={potentialJobs} isAdmin={isAdmin} />}
        {activeView === "smallJobs" && <SmallJobsPage isAdmin={isAdmin} smallJobs={smallJobs} setSmallJobs={setSmallJobs} />}
        {activeView === "inspections" && <InspectionsView jobs={visibleJobs} selectedJob={selectedJob} setSelectedId={setSelectedId} updateInspection={updateInspection} isAdmin={isAdmin} />}
        {activeView === "timeTracking" && <TimeTrackingView isAdmin={isAdmin} jobs={visibleJobs} currentUser={currentUser} />}
        {activeView === "labourTracking" && <LabourTrackingView isAdmin={isAdmin} jobs={visibleJobs} />}
      </main>
      <>
  {showSettings && (
    <SettingsModal
      users={users}
      setUsers={setUsers}
      currentUser={currentUser}
      onClose={() => setShowSettings(false)}
    />
  )}

  {showAddJob && (
    <AddJobModal
      users={users}
      onClose={() => setShowAddJob(false)}
      onCreate={(input) => {
        const today = new Date().toISOString().slice(0, 10);
        const newJob: Job = {
          id: Date.now(),
          jobNumber: `PJ-2026-${String(jobs.length + 1).padStart(3, "0")}`,
          name: input.name || "New Awarded Job",
          customer: input.customer || "New Customer",
          location: input.location || "Location",
          status: "Planning",
          certainty: "Confirmed",
          phase: "Planning",
          risk: "Green",
          startDate: input.startDate || today,
          finishDate: input.finishDate || today,
          nextAction: "Review project setup and assign first action item",
          actionOwner: currentUser.name,
          actionDue: input.startDate || today,
          progress: 0,
          allowedHours: input.allowedHours || 0,
          usedHours: 0,
          budget: input.budget || 0,
          costToDate: 0,
          labourBudget: input.labourBudget || 0,
          labourCostToDate: 0,
          crewIds: input.crewIds,
          crewTasks: [],
          projectTasks: pmTaskList(0),
          pmRequests: [],
          inspections: [],
        };
        setJobs([newJob, ...jobs]);
        setSelectedId(newJob.id);
        setShowAddJob(false);
        setActiveView("dashboard");
      }}
    />
  )}
</>
    </div>
  );
}

function LoginPage({ loginEmail, setLoginEmail, loginPassword, setLoginPassword, loginError, onLogin }: { loginEmail: string; setLoginEmail: (value: string) => void; loginPassword: string; setLoginPassword: (value: string) => void; loginError: string; onLogin: () => void }) {
  return (
    <div style={styles.loginShell}>
      <div style={styles.loginCard}>
        <div style={styles.loginLogo}>PJ'S<br />ELECTRIC</div>
        <div>
          <h1 style={styles.loginTitle}>JobFlow Login</h1>
          <p style={styles.muted}>Sign in to view your assigned jobs, daily reports, photos, inspections, crew points, and PM/Admin tools.</p>
        </div>

        <Field label="Email / Username">
          <input style={styles.input} type="email" value={loginEmail} onChange={(event) => setLoginEmail(event.target.value)} placeholder="name@pjselectric.ca" autoComplete="username" />
        </Field>

        <Field label="Password">
          <input style={styles.input} type="password" value={loginPassword} onChange={(event) => setLoginPassword(event.target.value)} placeholder="Prototype password field" onKeyDown={(event) => { if (event.key === "Enter") onLogin(); }} />
        </Field>

        {loginError && <div style={styles.loginError}>{loginError}</div>}
        <button style={styles.loginButton} onClick={onLogin}>Sign In</button>
        <p style={styles.loginNote}>Prototype login: Admin password is <strong>admin</strong>. Crew/Foreman temporary password is <strong>temp123</strong>, then they set their own password on first login. Online version should connect this to Supabase Auth.</p>
      </div>
    </div>
  );
}

function PasswordSetupPage({ user, onSave, onCancel }: { user: User; onSave: (password: string) => void; onCancel: () => void }) {
  const [password, setPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [error, setError] = useState("");

  function savePassword() {
    if (password.length < 6) {
      setError("Password must be at least 6 characters.");
      return;
    }
    if (password !== confirmPassword) {
      setError("Passwords do not match.");
      return;
    }
    onSave(password);
  }

  return (
    <div style={styles.loginShell}>
      <div style={styles.loginCard}>
        <div style={styles.loginLogo}>PJ'S<br />ELECTRIC</div>
        <div>
          <h1 style={styles.loginTitle}>Set Your Password</h1>
          <p style={styles.muted}>{user.name}, create your own password before opening JobFlow.</p>
        </div>
        <Field label="New Password">
          <input style={styles.input} type="password" value={password} onChange={(event) => setPassword(event.target.value)} placeholder="Minimum 6 characters" />
        </Field>
        <Field label="Confirm Password">
          <input style={styles.input} type="password" value={confirmPassword} onChange={(event) => setConfirmPassword(event.target.value)} onKeyDown={(event) => { if (event.key === "Enter") savePassword(); }} />
        </Field>
        {error && <div style={styles.loginError}>{error}</div>}
        <button style={styles.loginButton} onClick={savePassword}>Save Password & Sign In</button>
        <button style={styles.secondary} onClick={onCancel}>Back to Login</button>
      </div>
    </div>
  );
}

function Sidebar({ activeView, setActiveView, user, isAdmin, onAddJob, onOpenSettings, onSignOut }: { activeView: View; setActiveView: (view: View) => void; user: User; isAdmin: boolean; onAddJob: () => void; onOpenSettings: () => void; onSignOut: () => void }) {
  return (
    <aside style={styles.sidebar}>
      <div style={styles.logo}>PJ&apos;S<br />ELECTRIC</div>
      <h2 style={styles.sidebarTitle}>PJ&apos;S ELECTRIC</h2>
      <SideButton active={activeView === "dashboard"} icon={<LayoutDashboard size={18} />} label="Jobs" onClick={() => setActiveView("dashboard")} />
      <SideButton active={activeView === "calendar"} icon={<CalendarClock size={18} />} label="Calendar" onClick={() => setActiveView("calendar")} />
      {isAdmin && <SideButton active={activeView === "smallJobs"} icon={<ClipboardList size={18} />} label="Jobs To Do" onClick={() => setActiveView("smallJobs")} />}
      {isAdmin && <SideButton icon={<Plus size={18} />} label="Add Job" onClick={onAddJob} />}
      {isAdmin && <SideButton active={activeView === "potentialJobs"} icon={<BriefcaseBusiness size={18} />} label="Potential Jobs" onClick={() => setActiveView("potentialJobs")} />}
      <SideButton active={activeView === "crewPoints"} icon={<Users size={18} />} label="Crew Points" onClick={() => setActiveView("crewPoints")} />
      <SideButton icon={<Settings size={18} />} label="Settings" onClick={onOpenSettings} />

      <div style={styles.userCard}>
        <div style={styles.userTop}>
          <div style={styles.avatar}>{initials(user.name)}</div>
          <div style={styles.userText}>
            <strong>{user.name}</strong>
            <span>{user.role}</span>
            {user.role === "Crew" && <b>{user.points} pts</b>}
          </div>
        </div>
        <button style={styles.signOut} onClick={onSignOut}><LogOut size={16} /> Sign Out</button>
      </div>
    </aside>
  );
}

function Header({ currentUser, onOpenSettings }: { currentUser: User; onOpenSettings: () => void }) {
  return (
    <header style={styles.header}>
      <div>
        <div style={styles.brand}>PJ&apos;S ELECTRIC</div>
        <h1 style={styles.title}>JobFlow Operations Board</h1>
        <p style={styles.subtitle}>Planning, reports, crew tasks, crew points, inspections, time tracking, labour tracking, and PM workflows.</p>
      </div>
      <div style={styles.headerActionsLocked}>
        <span style={styles.signedIn}>Signed in as</span>
        <div style={styles.currentUserBadge}>
          <strong>{currentUser.name}</strong>
          <span>{currentUser.role}</span>
        </div>
        <button style={styles.secondary} onClick={onOpenSettings}><Settings size={16} /> Settings</button>
      </div>
    </header>
  );
}

function DashboardView({ jobs, selectedJob, selectedCrew, search, setSearch, setSelectedId, updateJob, deleteJob, toggleTask, updateInspection, isAdmin, currentUser, dailyReports, submitDailyReport, documents, uploadJobDocuments, potentialJobs, setPotentialJobs }: { jobs: Job[]; selectedJob: Job; selectedCrew: User[]; search: string; setSearch: (value: string) => void; setSelectedId: (id: number) => void; updateJob: (update: Partial<Job>) => void; deleteJob: (jobId: number) => void; toggleTask: (type: "crewTasks" | "projectTasks" | "pmRequests", id: number) => void; updateInspection: (id: number, status: string) => void; isAdmin: boolean; currentUser: User; dailyReports: DailyReport[]; submitDailyReport: (jobId: number, note: string, pictures?: string[], pictureUploads?: JobDocumentUpload[]) => void; documents: JobDocument[]; uploadJobDocuments: (jobId: number, uploads: JobDocumentUpload[]) => void; potentialJobs: PotentialJob[]; setPotentialJobs: (jobs: PotentialJob[]) => void }) {
  const [jobOpen, setJobOpen] = useState(false);
  const [activeFolder, setActiveFolder] = useState<JobFolder>("crewTasks");
  const [showPotentialForm, setShowPotentialForm] = useState(false);
  const [potentialForm, setPotentialForm] = useState({ name: "", customer: "", location: "", startDate: new Date().toISOString().slice(0, 10), finishDate: new Date().toISOString().slice(0, 10), estimatedHours: 0, assumedValue: 0, probability: "Medium" });

  function addPotentialJob() {
    if (!potentialForm.name.trim()) return;
    setPotentialJobs([
      {
        id: Date.now(),
        name: potentialForm.name.trim(),
        customer: potentialForm.customer || "TBD",
        location: potentialForm.location || "TBD",
        probability: potentialForm.probability,
        startDate: potentialForm.startDate,
        finishDate: potentialForm.finishDate,
        crewNeeded: 0,
        estimatedHours: potentialForm.estimatedHours,
        assumedValue: potentialForm.assumedValue,
        crewIds: [],
        document: "",
        documents: [],
        activityLog: [{ id: Date.now() + 99, date: new Date().toISOString().slice(0, 10), user: "System", field: "Original Entry", previousValue: "—", newValue: "Initial potential job created" }],
      },
      ...potentialJobs,
    ]);
    setPotentialForm({ name: "", customer: "", location: "", startDate: new Date().toISOString().slice(0, 10), finishDate: new Date().toISOString().slice(0, 10), estimatedHours: 0, assumedValue: 0, probability: "Medium" });
    setShowPotentialForm(false);
  }

  if (!jobOpen) {
    return (
      <div style={styles.cleanStack}>
        <PageTitle title="Jobs" subtitle={isAdmin ? "Click a job to open the full job file." : "Crew and foremen only see the jobs they are assigned to."} />
        <Card>
          <div style={styles.cardHeader}>
            <div>
              <h2 style={styles.cardTitle}>Assigned / Active Jobs</h2>
              <p style={styles.muted}>{isAdmin ? "Admin/PM can see all jobs." : "Your view is limited to assigned jobs only."}</p>
            </div>
            <SearchBox value={search} onChange={setSearch} />
          </div>
          <JobList jobs={jobs} selectedJob={selectedJob} setSelectedId={(id) => { setSelectedId(id); setActiveFolder("crewTasks"); setJobOpen(true); }} />
        </Card>

        </div>
    );
  }

  return (
    <div style={styles.cleanStack}>
      <div style={styles.jobFileHeader}>
        <button style={styles.secondary} onClick={() => setJobOpen(false)}>← Back to Jobs</button>
        <div>
          <h2 style={styles.jobFileTitle}>{selectedJob.name}</h2>
          <p style={styles.jobDetailSub}>{selectedJob.jobNumber} · {selectedJob.customer} · {selectedJob.location}</p>
        </div>
        {isAdmin && <button type="button" style={styles.danger} onClick={(event) => { event.preventDefault(); event.stopPropagation(); deleteJob(selectedJob.id); setJobOpen(false); }}><Trash2 size={15} /> Delete Job</button>}
      </div>

      <HealthSummary job={selectedJob} crew={selectedCrew} isAdmin={isAdmin} />

      <section style={styles.jobFileMainGrid}>
        <Card>
          <div style={styles.cardHeader}>
            <h2 style={styles.cardTitle}>Job Main Screen</h2>
            <Badge tone={selectedJob.risk === "Green" ? "green" : selectedJob.risk === "Red" ? "red" : "yellow"}>{selectedJob.risk}</Badge>
          </div>
          <DetailGrid job={selectedJob} />
          {isAdmin && <AdminProjectStatusEdit job={selectedJob} updateJob={updateJob} />}
          <StatusSnapshot job={selectedJob} />
          <HealthBars job={selectedJob} isAdmin={isAdmin} />
          <NextAction job={selectedJob} updateJob={updateJob} />
        </Card>

        <Card>
          <h2 style={styles.cardTitle}>Job Folders</h2>
          <JobFolderTabs activeFolder={activeFolder} setActiveFolder={setActiveFolder} />
          <JobFolderContent
            job={selectedJob}
            activeFolder={activeFolder}
            isAdmin={isAdmin}
            updateJob={updateJob}
            updateInspection={updateInspection}
            toggleTask={toggleTask}
            dailyReports={dailyReports}
            currentUser={currentUser}
            submitDailyReport={submitDailyReport}
            documents={documents}
            uploadJobDocuments={uploadJobDocuments}
          />
        </Card>
      </section>
    </div>
  );
}

function AdminProjectStatusEdit({ job, updateJob }: { job: Job; updateJob: (update: Partial<Job>) => void }) {
  return (
    <div style={styles.statusEditBox}>
      <h3 style={styles.sectionTitle}>Edit Project Status</h3>
      <div style={styles.statusEditGrid}>
        <Field label="Project Status"><select style={styles.input} value={job.status} onChange={(event) => updateJob({ status: event.target.value as JobStatus })}><option>Planning</option><option>Active</option><option>On Hold</option><option>Complete</option></select></Field>
        <Field label="Current Phase"><input style={styles.input} value={job.phase} onChange={(event) => updateJob({ phase: event.target.value })} /></Field>
        <Field label="Risk"><select style={styles.input} value={job.risk} onChange={(event) => updateJob({ risk: event.target.value as Risk })}><option>Green</option><option>Yellow</option><option>Red</option></select></Field>
        <Field label="Progress %"><input style={styles.input} type="number" min={0} max={100} value={job.progress} onChange={(event) => updateJob({ progress: Number(event.target.value) })} /></Field>
        <Field label="Action Owner"><input style={styles.input} value={job.actionOwner} onChange={(event) => updateJob({ actionOwner: event.target.value })} /></Field>
        <Field label="Action Due"><input style={styles.input} type="date" value={job.actionDue} onChange={(event) => updateJob({ actionDue: event.target.value })} /></Field>
      </div>
    </div>
  );
}

function JobFolderTabs({ activeFolder, setActiveFolder }: { activeFolder: JobFolder; setActiveFolder: (folder: JobFolder) => void }) {
  const folders: { id: JobFolder; label: string }[] = [
    { id: "crewTasks", label: "Crew Tasks" },
    { id: "inspections", label: "Inspections" },
    { id: "pmTasks", label: "PM Tasks" },
    { id: "dailyNotes", label: "Daily Notes / Pictures" },
    { id: "photos", label: "Photos" },
    { id: "documents", label: "Documents" },
    { id: "knowifyUploads", label: "Knowify Uploads" },
    { id: "activityLog", label: "Activity Log" },
  ];
  return <div style={styles.jobFolderTabs}>{folders.map((folder) => <button key={folder.id} style={activeFolder === folder.id ? styles.folderTabActive : styles.folderTab} onClick={() => setActiveFolder(folder.id)}>{folder.label}</button>)}</div>;
}

function JobFolderContent({ job, activeFolder, isAdmin, updateJob, updateInspection, toggleTask, dailyReports, currentUser, submitDailyReport, documents, uploadJobDocuments }: { job: Job; activeFolder: JobFolder; isAdmin: boolean; updateJob: (update: Partial<Job>) => void; updateInspection: (id: number, status: string) => void; toggleTask: (type: "crewTasks" | "projectTasks" | "pmRequests", id: number) => void; dailyReports: DailyReport[]; currentUser: User; submitDailyReport: (jobId: number, note: string, pictures?: string[], pictureUploads?: JobDocumentUpload[]) => void; documents: JobDocument[]; uploadJobDocuments: (jobId: number, uploads: JobDocumentUpload[]) => void }) {
  const [newCrewTask, setNewCrewTask] = useState("");
  const [newInspectionTitle, setNewInspectionTitle] = useState("");
  const [newPmTaskTitle, setNewPmTaskTitle] = useState("");
  const [activityDate, setActivityDate] = useState(new Date().toISOString().slice(0, 10));
  const [showKnowifyUpload, setShowKnowifyUpload] = useState(false);
  const [documentFiles, setDocumentFiles] = useState<JobDocumentUpload[]>([]);
  const [knowifyUploadType, setKnowifyUploadType] = useState<"P&L Report" | "Time Report">("Time Report");
  const [knowifyFiles, setKnowifyFiles] = useState<string[]>([]);
  const [isParsingKnowify, setIsParsingKnowify] = useState(false);
  const [importedHours, setImportedHours] = useState(job.usedHours);
  const [importedLabourCost, setImportedLabourCost] = useState(job.labourCostToDate || 0);
  const [importedBudgetLeft, setImportedBudgetLeft] = useState(Math.max(0, job.budget - job.costToDate));
  const [lastKnowifyUpdate, setLastKnowifyUpdate] = useState("");

  const labourUsedPct = pct(job.usedHours, job.allowedHours);
  const labourLeftPct = Math.max(0, 100 - labourUsedPct);
  const budgetUsedPct = pct(job.costToDate, job.budget);
  const budgetLeft = Math.max(0, job.budget - job.costToDate);
  const budgetHealthPct = Math.max(0, 100 - budgetUsedPct);

  function parseHourValue(value: unknown): number {
    if (typeof value === "number") return value;
    if (typeof value !== "string") return 0;
    const trimmed = value.trim();
    if (!trimmed) return 0;
    if (trimmed.includes(":")) {
      const [hours, minutes] = trimmed.split(":").map((part) => Number(part));
      return (hours || 0) + (minutes || 0) / 60;
    }
    const numeric = Number(trimmed);
    return Number.isFinite(numeric) ? numeric : 0;
  }

  function parseMoneyValue(value: unknown): number {
    if (typeof value === "number") return value;
    if (typeof value !== "string") return 0;
    const cleaned = value.replace(/[$,()]/g, "").trim();
    const numeric = Number(cleaned);
    return Number.isFinite(numeric) ? Math.abs(numeric) : 0;
  }

  async function parseKnowifyFile(file: File) {
    setIsParsingKnowify(true);
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array", cellDates: false });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<any[]>(firstSheet, { header: 1, defval: null });

      if (knowifyUploadType === "Time Report") {
        const totalRow = rows.find((row) => String(row?.[0] || "").toLowerCase().trim() === "total");
        const totalHours = parseHourValue(totalRow?.[3] ?? totalRow?.[2]);
        const totalLabourCost = parseMoneyValue(totalRow?.[4]);
        if (totalHours > 0 || totalLabourCost > 0) {
          const rounded = Number(totalHours.toFixed(2));
          setImportedHours(rounded);
          setImportedLabourCost(totalLabourCost);
          updateJob({ usedHours: rounded, labourCostToDate: totalLabourCost });
          setLastKnowifyUpdate(`Time Report applied: ${rounded.toLocaleString()} hours and ${currency(totalLabourCost)} labour cost to date. Labour budget health recalculated automatically.`);
        } else {
          setLastKnowifyUpdate("Time Report uploaded, but total hours could not be found. Enter Hours Used To Date manually, then save.");
        }
      } else {
        const totalCostsRow = rows.find((row) => String(row?.[0] || "").toLowerCase().trim() === "total costs");
        const currentCost = parseMoneyValue(totalCostsRow?.[2]);
        const varianceBudgetLeft = parseMoneyValue(totalCostsRow?.[3]);
        if (currentCost > 0) {
          const budgetLeft = Number(Math.max(0, job.budget - currentCost).toFixed(2));
          setImportedBudgetLeft(budgetLeft);
          updateJob({ costToDate: currentCost });
          setLastKnowifyUpdate(`P&L Report applied: ${currency(currentCost)} current costs from the Total Costs row. Budget left = ${currency(budgetLeft)}. Budget health recalculated automatically.`);
        } else if (varianceBudgetLeft > 0) {
          const costToDate = Math.max(0, job.budget - varianceBudgetLeft);
          setImportedBudgetLeft(varianceBudgetLeft);
          updateJob({ costToDate });
          setLastKnowifyUpdate(`P&L Report applied: ${currency(varianceBudgetLeft)} budget left from Knowify variance. Budget health recalculated automatically.`);
        } else {
          setLastKnowifyUpdate("P&L Report uploaded, but Total Costs variance could not be found. Enter Variance / Budget Left manually, then save.");
        }
      }
    } catch (error) {
      setLastKnowifyUpdate("Could not read the Knowify Excel file in the browser. Enter the report total manually, then save.");
    } finally {
      setIsParsingKnowify(false);
    }
  }

  function saveKnowifyUpload() {
    if (knowifyUploadType === "Time Report") {
      const hoursToDate = Math.max(0, Number(importedHours) || 0);
      const labourCostToDate = Math.max(0, Number(importedLabourCost) || 0);
      updateJob({ usedHours: hoursToDate, labourCostToDate });
      setImportedHours(hoursToDate);
      setImportedLabourCost(labourCostToDate);
      setLastKnowifyUpdate(`Time Report applied: ${hoursToDate.toLocaleString()} used hours and ${currency(labourCostToDate)} labour cost to date. Labour budget used = ${pct(labourCostToDate, job.labourBudget || 0)}% of ${currency(job.labourBudget || 0)}.`);
    } else {
      const positiveBudgetLeft = Math.max(0, Math.abs(Number(importedBudgetLeft) || 0));
      const costToDate = Math.max(0, job.budget - positiveBudgetLeft);
      updateJob({ costToDate });
      setImportedBudgetLeft(positiveBudgetLeft);
      setLastKnowifyUpdate(`P&L Report applied: ${positiveBudgetLeft.toLocaleString()} budget left. Cost to date = ${costToDate.toLocaleString()}. Budget health = ${Math.max(0, 100 - pct(costToDate, job.budget))}%.`);
    }
    setShowKnowifyUpload(false);
  }
  const [selectedDate, setSelectedDate] = useState(new Date().toISOString().slice(0, 10));
  const [dailyNote, setDailyNote] = useState("");
  const [pictureNames, setPictureNames] = useState<string[]>([]);
  const [pictureUploads, setPictureUploads] = useState<JobDocumentUpload[]>([]);
  const [previewDocument, setPreviewDocument] = useState<JobDocument | null>(null);
  const jobReports = dailyReports.filter((report) => report.jobId === job.id);
  const visibleReports = isAdmin ? jobReports : jobReports.filter((report) => report.userId === currentUser.id);
  const filteredReports = isAdmin ? jobReports.filter((report) => report.date === selectedDate) : visibleReports;
  const canSubmitDailyNotes = true;
  const earnsDailyReportPoints = currentUser.role === "Crew" || currentUser.role === "Foreman";
  const today = new Date().toISOString().slice(0, 10);
  const existingTodayReport = visibleReports.find((report) => report.userId === currentUser.id && report.date === today);

  if (activeFolder === "crewTasks") {
    return (
      <div style={styles.folderPanel}>
        <div style={styles.addInlineBox}>
          <h3 style={styles.sideTitle}>Add Crew Task</h3>
          <div style={styles.pmRequestForm}>
            <input style={styles.input} value={newCrewTask} onChange={(event) => setNewCrewTask(event.target.value)} placeholder="Add a crew task for this job..." />
            <button style={styles.smallButton} onClick={() => { if (!newCrewTask.trim()) return; updateJob({ crewTasks: [{ id: Date.now(), title: `${newCrewTask.trim()} — added by ${currentUser.name}`, done: false }, ...job.crewTasks] }); setNewCrewTask(""); }}><Plus size={16} /> Add Task</button>
          </div>
        </div>
        <TaskCard title="Crew Tasks" tasks={job.crewTasks} onToggle={(id) => toggleTask("crewTasks", id)} canEdit />
      </div>
    );
  }

  if (activeFolder === "inspections") {
    return (
      <div style={styles.folderPanel}>
        {isAdmin && (
          <div style={styles.addInlineBox}>
            <h3 style={styles.sideTitle}>Add Inspection</h3>
            <div style={styles.pmRequestForm}>
              <input style={styles.input} value={newInspectionTitle} onChange={(event) => setNewInspectionTitle(event.target.value)} placeholder="Example: Rough-in inspection, Final inspection..." />
              <button style={styles.smallButton} onClick={() => { if (!newInspectionTitle.trim()) return; updateJob({ inspections: [{ id: Date.now(), title: newInspectionTitle.trim(), status: "Not Started", date: "", notes: "" }, ...job.inspections] }); setNewInspectionTitle(""); }}><Plus size={16} /> Add Inspection</button>
            </div>
          </div>
        )}
        <InspectionCard inspections={job.inspections} canEdit={isAdmin} onChange={updateInspection} />
      </div>
    );
  }

  if (activeFolder === "pmTasks") {
    return (
      <div style={styles.folderPanel}>
        {isAdmin && (
          <div style={styles.addInlineBox}>
            <h3 style={styles.sideTitle}>Add PM Task</h3>
            <p style={styles.cardHelp}>Add a task to the PM checklist for this job. This is added on top of the preloaded checklist.</p>
            <div style={styles.pmRequestForm}>
              <input style={styles.input} value={newPmTaskTitle} onChange={(event) => setNewPmTaskTitle(event.target.value)} placeholder="Example: Submit revised shop drawings, price PCN, confirm inspection..." />
              <button style={styles.smallButton} onClick={() => { if (!newPmTaskTitle.trim()) return; updateJob({ projectTasks: [{ id: Date.now(), title: `${newPmTaskTitle.trim()} — added by ${currentUser.name}`, done: false }, ...job.projectTasks] }); setNewPmTaskTitle(""); }}><Plus size={16} /> Add PM Task</button>
            </div>
          </div>
        )}
        {isAdmin ? (
          <>
            <TaskCard title="Project Management Task List" tasks={job.projectTasks} onToggle={(id) => toggleTask("projectTasks", id)} canEdit lockedNote="PM/Admin internal checklist." />
            <PMRequestCard job={job} updateJob={updateJob} toggleTask={(id) => toggleTask("pmRequests", id)} isAdmin={isAdmin} currentUser={currentUser} />
          </>
        ) : (
          <>
            <Card>
              <h3 style={styles.sideTitle}>PM Requests</h3>
              <p style={styles.muted}>Create tasks here for the PM/Admin to complete. You can submit the request, but only PM/Admin can check it off when it is done.</p>
            </Card>
            <PMRequestCard job={job} updateJob={updateJob} toggleTask={(id) => toggleTask("pmRequests", id)} isAdmin={isAdmin} currentUser={currentUser} />
          </>
        )}
      </div>
    );
  }

  if (activeFolder === "activityLog") {
    if (!isAdmin) {
      return (
        <div style={styles.folderPanel}>
          <Card>
            <h3 style={styles.sideTitle}>Activity Log</h3>
            <p style={styles.muted}>Only Admin/PM can view the full job activity log.</p>
          </Card>
        </div>
      );
    }

    const dailyItems = dailyReports
      .filter((report) => report.jobId === job.id && report.date === activityDate)
      .map((report) => ({ type: "Daily Report", title: `${report.userName} submitted daily notes`, detail: `${report.note} (+${report.pointsAwarded} pts)` }));

    // Activity log is date-based. Only items with the selected activityDate are shown here.
    // Later, when this is connected to the database, every upload/task/status change should save its own date.
    const uploadItems: { type: string; title: string; detail: string }[] = documents
      .filter((document) => document.jobId === job.id && document.uploadedDate === activityDate)
      .map((document) => ({
        type: document.type === "Photo" ? "Photo Upload" : "Document Upload",
        title: `${document.uploadedBy} uploaded ${document.fileName}`,
        detail: `${document.type} added to the job file`,
      }));
    const crewTaskItems: { type: string; title: string; detail: string }[] = [];
    const pmTaskItems: { type: string; title: string; detail: string }[] = [];
    const inspectionItems: { type: string; title: string; detail: string }[] = [];
    const adminItems: { type: string; title: string; detail: string }[] = [];

    const activityItems = [...dailyItems, ...uploadItems, ...crewTaskItems, ...pmTaskItems, ...inspectionItems, ...adminItems];

    return (
      <div style={styles.folderPanel}>
        <div style={styles.compactCardHeader}>
          <div>
            <h3 style={styles.sectionTitle}>Activity Log</h3>
            <p style={styles.cardHelp}>PM/Admin calendar view of what happened on the selected day: daily notes, uploads, task completions, inspection updates, and admin/job changes.</p>
          </div>
          <input style={styles.input} type="date" value={activityDate} onChange={(event) => setActivityDate(event.target.value)} />
        </div>
        <div style={styles.activityList}>
          {activityItems.length === 0 ? (
            <div style={styles.reportHistoryItem}>
              <strong>No activity for {activityDate}</strong>
              <span>No daily notes, uploads, task updates, or admin changes found for this date.</span>
            </div>
          ) : (
            activityItems.map((item, index) => (
              <div style={styles.activityItem} key={`${item.type}-${index}`}>
                <span style={styles.activityType}>{item.type}</span>
                <strong>{item.title}</strong>
                <p>{item.detail}</p>
              </div>
            ))
          )}
        </div>
      </div>
    );
  }

  if (activeFolder === "photos") {
    const jobPhotos = documents.filter((document) => document.jobId === job.id && document.type === "Photo");
    const groupedPhotos = jobPhotos.reduce<Record<string, JobDocument[]>>((groups, photo) => {
      groups[photo.uploadedDate] = groups[photo.uploadedDate] || [];
      groups[photo.uploadedDate].push(photo);
      return groups;
    }, {});
    const photoDates = Object.keys(groupedPhotos).sort().reverse();

    return (
      <div style={styles.folderPanel}>
        <div style={styles.compactCardHeader}>
          <div>
            <h3 style={styles.sectionTitle}>Job Photos</h3>
            <p style={styles.cardHelp}>All photos uploaded to this job show here, including photos attached to daily reports for the +5 photo bonus.</p>
          </div>
          <span style={styles.smallPill}>{jobPhotos.length} photos</span>
        </div>

        <div style={styles.knowifyUploadPanel}>
          <h3 style={styles.sideTitle}>Upload Job Photos</h3>
          <label style={styles.dropZone}>
            <strong>Click to upload photos</strong>
            <span>Crew, foremen, PM, and admin can add job photos here.</span>
            <input
              style={styles.hiddenFileInput}
              type="file"
              multiple
              accept="image/*"
              onClick={(event) => {
                (event.target as HTMLInputElement).value = "";
              }}
              onChange={(event) => {
                const files = Array.from(event.target.files || []);
                setDocumentFiles(files.map((file) => ({
                  fileName: file.name,
                  type: "Photo" as const,
                  url: URL.createObjectURL(file),
                })));
              }}
            />
          </label>
          {documentFiles.length > 0 && <div style={styles.pictureList}>{documentFiles.map((file) => <span key={file.fileName}>{file.fileName}</span>)}</div>}
          <button style={styles.primary} onClick={() => { uploadJobDocuments(job.id, documentFiles.filter((file) => file.type === "Photo")); setDocumentFiles([]); }}>Save Photos</button>
        </div>

        <div style={styles.documentFolderList}>
          {photoDates.length === 0 && <div style={styles.reportHistoryItem}><strong>No photos uploaded yet</strong><span>Photos from daily reports and the photo uploader will show here.</span></div>}
          {photoDates.map((date) => (
            <div style={styles.documentDateFolder} key={date}>
              <div style={styles.compactCardHeader}>
                <h3 style={styles.sideTitle}>{date}</h3>
                <span style={styles.smallPill}>{groupedPhotos[date].length} photos</span>
              </div>
              <div style={styles.documentGrid}>
                {groupedPhotos[date].map((photo) => (
                  <button style={styles.documentCard} key={photo.id} onClick={() => setPreviewDocument(photo)}>
                    {photo.url ? <img src={photo.url} alt={photo.fileName} style={styles.documentThumb} /> : <div style={styles.documentIcon}>📷</div>}
                    <strong>{photo.fileName}</strong>
                    <span>Uploaded by {photo.uploadedBy}</span>
                    <small>{photo.uploadedAt || photo.uploadedDate}</small>
                  </button>
                ))}
              </div>
            </div>
          ))}
        </div>
        {previewDocument && <DocumentPreviewModal document={previewDocument} onClose={() => setPreviewDocument(null)} />}
      </div>
    );
  }

  if (activeFolder === "documents") {
    const jobDocuments = documents.filter((document) => document.jobId === job.id && document.type === "Document");
    const groupedDocuments = jobDocuments.reduce<Record<string, JobDocument[]>>((groups, document) => {
      groups[document.uploadedDate] = groups[document.uploadedDate] || [];
      groups[document.uploadedDate].push(document);
      return groups;
    }, {});
    const documentDates = Object.keys(groupedDocuments).sort().reverse();

    return (
      <div style={styles.folderPanel}>
        <div style={styles.compactCardHeader}>
          <div>
            <h3 style={styles.sectionTitle}>Documents</h3>
            <p style={styles.cardHelp}>PM/Admin document folder for this job. Daily notes stay in the Daily Notes tab and photos stay in the Photos tab.</p>
          </div>
          <span style={styles.smallPill}>{jobDocuments.length} documents</span>
        </div>

        {isAdmin && (
          <div style={styles.knowifyUploadPanel}>
            <h3 style={styles.sideTitle}>Upload PM/Admin Documents</h3>
            <label style={styles.dropZone}>
              <strong>Click to upload job documents</strong>
              <span>Drawings, PDFs, Excel files, reports, site instructions, closeout files, etc.</span>
              <input
                style={styles.hiddenFileInput}
                type="file"
                multiple
                onClick={(event) => {
                  (event.target as HTMLInputElement).value = "";
                }}
                onChange={(event) => {
                  const files = Array.from(event.target.files || []);
                  setDocumentFiles(files.map((file) => ({
                    fileName: file.name,
                    type: "Document" as const,
                    url: URL.createObjectURL(file),
                  })));
                }}
              />
            </label>
            {documentFiles.length > 0 && <div style={styles.pictureList}>{documentFiles.map((file) => <span key={file.fileName}>{file.fileName}</span>)}</div>}
            <button style={styles.primary} onClick={() => { uploadJobDocuments(job.id, documentFiles.filter((file) => file.type === "Document")); setDocumentFiles([]); }}>Save Documents</button>
          </div>
        )}

        {!isAdmin && <div style={styles.reportHistoryItem}><strong>View Only</strong><span>Only PM/Admin can upload documents. Crew/foremen can view documents that PM/Admin upload here.</span></div>}

        <div style={styles.documentFolderList}>
          {documentDates.length === 0 && <div style={styles.reportHistoryItem}><strong>No documents uploaded yet</strong><span>PM/Admin uploaded documents will show here by upload date.</span></div>}
          {documentDates.map((date) => (
            <div style={styles.documentDateFolder} key={date}>
              <div style={styles.compactCardHeader}>
                <h3 style={styles.sideTitle}>{date}</h3>
                <span style={styles.smallPill}>{groupedDocuments[date].length} documents</span>
              </div>
              <div style={styles.documentGrid}>
                {groupedDocuments[date].map((document) => (
                  <button style={styles.documentCard} key={document.id} onClick={() => setPreviewDocument(document)}>
                    <div style={styles.documentIcon}>📄</div>
                    <strong>{document.fileName}</strong>
                    <span>Uploaded by {document.uploadedBy}</span>
                    <small>{document.uploadedAt || document.uploadedDate}</small>
                  </button>
                ))}
              </div>
            </div>
          ))}
        </div>
        {previewDocument && <DocumentPreviewModal document={previewDocument} onClose={() => setPreviewDocument(null)} />}
      </div>
    );
  }

  if (activeFolder === "knowifyUploads") {
    if (!isAdmin) {
      return (
        <div style={styles.folderPanel}>
          <Card>
            <h3 style={styles.sideTitle}>Knowify Uploads</h3>
            <p style={styles.muted}>Only Admin/PM can upload Knowify P&L reports and time reports.</p>
          </Card>
        </div>
      );
    }

    return (
      <div style={styles.folderPanel}>
        <div style={styles.compactCardHeader}>
          <div>
            <h3 style={styles.sectionTitle}>Knowify Uploads</h3>
            <p style={styles.cardHelp}>Upload Knowify P&L reports and time reports for this selected job. For this prototype, enter the totals from the report below: Time Report = hours used to date. P&L Report = reads the Total Costs row Current results as cost to date. If that is not found, it falls back to Knowify variance/budget left and converts it positive if exported as negative.</p>
          </div>
          <button style={styles.primary} onClick={() => setShowKnowifyUpload(true)}><Plus size={16} /> Upload Report</button>
        </div>

        <div style={styles.uploadStatsGrid}>
          <div style={styles.uploadStat}><strong>Time Report</strong><span>{job.usedHours.toLocaleString()} hrs used to date</span></div>
          <div style={styles.uploadStat}><strong>P&L Variance</strong><span>{budgetLeft.toLocaleString()} budget left</span></div>
          <div style={styles.uploadStat}><strong>Labour Health</strong><span>{labourLeftPct}% left / {labourUsedPct}% used</span></div>
          <div style={styles.uploadStat}><strong>Budget Health</strong><span>{budgetHealthPct}% left / {budgetUsedPct}% used</span></div>
        </div>
        {lastKnowifyUpdate && <div style={styles.successNotice}>{lastKnowifyUpdate}</div>}

        {showKnowifyUpload && (
          <div style={styles.knowifyUploadPanel}>
            <div style={styles.compactCardHeader}>
              <h3 style={styles.sideTitle}>Upload Knowify Report</h3>
              <button style={styles.closeMiniButton} onClick={() => setShowKnowifyUpload(false)}>×</button>
            </div>
            <div style={styles.uploadFormGrid}>
              <Field label="Report Type"><select style={styles.input} value={knowifyUploadType} onChange={(event) => setKnowifyUploadType(event.target.value as "P&L Report" | "Time Report")}><option>Time Report</option><option>P&L Report</option></select></Field>
              <Field label="Job"><input style={styles.input} value={job.name} readOnly /></Field>
              <Field label="Start Date"><input style={styles.input} type="date" /></Field>
              <Field label="End Date"><input style={styles.input} type="date" /></Field>
              {knowifyUploadType === "Time Report" && <Field label="Hours Used To Date"><input style={styles.input} type="number" value={importedHours} onChange={(event) => setImportedHours(Number(event.target.value))} placeholder="Total hours to date" /></Field>}
              {knowifyUploadType === "Time Report" && <Field label="Labour Cost To Date"><input style={styles.input} type="number" value={importedLabourCost} onChange={(event) => setImportedLabourCost(Number(event.target.value))} placeholder="Total labour cost to date" /></Field>}
              {knowifyUploadType === "P&L Report" && <Field label="Budget Left / Variance"><input style={styles.input} type="number" value={importedBudgetLeft} onChange={(event) => setImportedBudgetLeft(Number(event.target.value))} placeholder="Budget remaining if entering manually" /></Field>}
            </div>
            <label style={styles.dropZone}>
              <strong>Drag and drop Knowify file here</strong>
              <span>or click to select Excel / CSV / PDF report files</span>
              <input style={styles.hiddenFileInput} type="file" multiple accept=".xlsx,.xls,.csv,.pdf" onChange={(event) => {
                const files = Array.from(event.target.files || []);
                setKnowifyFiles(files.map((file) => file.name));
                if (files[0]) parseKnowifyFile(files[0]);
              }} />
            </label>
            {knowifyFiles.length > 0 && <div style={styles.pictureList}>{knowifyFiles.map((file) => <span key={file}>{file}</span>)}</div>}
            <button style={styles.primary} onClick={saveKnowifyUpload}>Save Upload / Recalculate Health</button>
          </div>
        )}
      </div>
    );
  }

  const uniqueDates = Array.from(new Set(visibleReports.map((report) => report.date))).sort().reverse();
  return (
    <div style={styles.folderPanel}>
      <div style={styles.compactCardHeader}>
        <div>
          <h3 style={styles.sectionTitle}>{isAdmin ? "Daily Notes Calendar" : "Daily Notes / Pictures"}</h3>
          <p style={styles.cardHelp}>{isAdmin ? "Choose a calendar date to show every daily note submitted on that day for this job, grouped by each crew member/foreman." : "Crew/foremen can add daily notes, picture names, and view their submitted notes for this job."}</p>
        </div>
        {isAdmin && <input style={styles.input} type="date" value={selectedDate} onChange={(event) => setSelectedDate(event.target.value)} />}
      </div>

      {isAdmin && (
        <div style={styles.calendarDateGrid}>
          {uniqueDates.length === 0 && <span style={styles.calendarEmpty}>No daily notes submitted yet</span>}
          {uniqueDates.map((date) => <button key={date} style={selectedDate === date ? styles.calendarDateActive : styles.calendarDate} onClick={() => setSelectedDate(date)}>{date}</button>)}
        </div>
      )}

      {canSubmitDailyNotes && (
        <div style={styles.dailyNoteComposer}>
          <div style={styles.compactCardHeader}>
            <div>
              <h3 style={styles.sideTitle}>Add Daily Note</h3>
              <p style={styles.cardHelp}>{earnsDailyReportPoints ? (existingTodayReport ? "You already submitted today's daily report. Editing it will not add the normal +5 again. Upload 5 or more photos with the daily report to earn one additional +5 photo bonus for today." : "Submit your daily notes here. First daily note today earns +5 points. Upload 5 or more photos with the daily report to earn an additional +5 photo bonus for today.") : "Add PM/Admin notes and attach photos for this job."}</p>
            </div>
            {earnsDailyReportPoints && <span style={styles.smallPill}>{existingTodayReport ? "Already earned today" : "+5 pts"}</span>}
          </div>
          <textarea style={styles.dailyReportInput} value={dailyNote || existingTodayReport?.note || ""} onChange={(event) => setDailyNote(event.target.value)} placeholder="Work completed, delays, materials needed, safety notes, inspections, or issues..." />
          <label style={styles.dropZone}>
            <strong>Click here to upload daily report photos</strong>
            <span>{earnsDailyReportPoints ? "Upload 5 or more photos with this daily report to earn an additional +5 points today." : "Upload photos for this daily note."}</span>
            <input
              style={styles.hiddenFileInput}
              type="file"
              multiple
              accept="image/*"
              onClick={(event) => {
                (event.target as HTMLInputElement).value = "";
              }}
              onChange={(event) => {
                const files = Array.from(event.target.files || []);
                const uploads = files.map((file) => ({
                  fileName: file.name,
                  type: "Photo" as const,
                  url: URL.createObjectURL(file),
                }));
                setPictureUploads(uploads);
                setPictureNames(uploads.map((file) => file.fileName));
              }}
            />
          </label>
          {earnsDailyReportPoints && pictureNames.length >= 5 ? (
  <div style={styles.successNotice}>
    Photo bonus ready: +5 additional points will be added for uploading 5 or more photos today.
  </div>
) : null}
          {pictureNames.length > 0 && <div style={styles.pictureList}>{pictureNames.map((name) => <span key={name}>{name}</span>)}</div>}
          <button style={styles.primary} onClick={() => { submitDailyReport(job.id, dailyNote || existingTodayReport?.note || "", pictureNames, pictureUploads); setDailyNote(""); setPictureNames([]); setPictureUploads([]); }}>{existingTodayReport ? "Save Daily Report Edit" : "Submit Daily Note"}</button>
        </div>
      )}

      <div style={styles.reportHistoryList}>
        {filteredReports.length === 0 && <div style={styles.reportHistoryItem}><strong>No notes for {selectedDate}</strong><span>Choose another date from the calendar above, or wait for crew/foremen to submit daily notes for this job.</span></div>}
        {filteredReports.sort((a, b) => a.userName.localeCompare(b.userName) || b.id - a.id).map((report) => {
          const reportPictures = report.pictures.length ? report.pictures : (report.pictureUploads || []).map((upload) => upload.fileName);
          const reportDocuments = documents.filter((item) => item.jobId === report.jobId && item.uploadedDate === report.date && item.uploadedBy === report.userName && item.type === "Photo");
          const combinedPictures = Array.from(new Set([...reportPictures, ...reportDocuments.map((document) => document.fileName)]));

          return (
            <div style={styles.reportHistoryItem} key={report.id}>
              <strong>{report.date} · {report.userName} · +{report.pointsAwarded} pts</strong>
              <span>{report.note}</span>
              <div style={styles.picturePlaceholder}>
                {combinedPictures.length ? combinedPictures.map((picture) => {
                  const upload = report.pictureUploads?.find((item) => item.fileName === picture);
                  const documentMatch = reportDocuments.find((item) => item.fileName === picture) || documents.find((item) => item.jobId === report.jobId && item.fileName === picture && item.uploadedDate === report.date && item.uploadedBy === report.userName);
                  const previewItem: JobDocument = documentMatch || {
                    id: report.id,
                    jobId: report.jobId,
                    fileName: picture,
                    uploadedBy: report.userName,
                    uploadedDate: report.date,
                    type: upload?.type || "Photo",
                    url: upload?.url,
                  };
                  return (
                    <button key={picture} style={styles.picturePreviewButton} onClick={() => setPreviewDocument(previewItem)}>
                      {previewItem.url ? "📷 View" : "📎 File"} {picture}
                    </button>
                  );
                }) : "No pictures attached"}
              </div>
            </div>
          );
        })}
        {previewDocument && <DocumentPreviewModal document={previewDocument} onClose={() => setPreviewDocument(null)} />}
      </div>
    </div>
  );
}

function HealthSummary({ job, crew, isAdmin }: { job: Job; crew: User[]; isAdmin: boolean }) {
  const labourUsed = pct(job.labourCostToDate || 0, job.labourBudget || 0);
  const budgetUsed = pct(job.costToDate, job.budget);
  return (
    <section style={styles.healthTiles}>
      <HealthTile title="Job Progress" value={`${job.progress}%`} subtitle="Overall completion" percent={job.progress} />
      <HealthTile title="Labour Used" value={isAdmin ? `${currency(job.labourCostToDate || 0)} / ${currency(job.labourBudget || 0)}` : `${labourUsed}%`} subtitle={isAdmin ? `${labourUsed}% of labour budget` : "Of labour budget"} percent={labourUsed} warn={labourUsed > 85} />
      <HealthTile title="Labour Left" value={isAdmin ? currency(Math.max(0, (job.labourBudget || 0) - (job.labourCostToDate || 0))) : `${100 - labourUsed}%`} subtitle="Remaining labour budget" percent={100 - labourUsed} blue />
      <HealthTile title="Budget Health" value={isAdmin ? currency(Math.max(0, job.budget - job.costToDate)) : `${100 - budgetUsed}%`} subtitle={isAdmin ? `${budgetUsed}% budget used` : "Budget remaining"} percent={100 - budgetUsed} blue warn={budgetUsed > 85} />
      <CrewCard crew={crew} />
    </section>
  );
}

function JobList({ jobs, selectedJob, setSelectedId }: { jobs: Job[]; selectedJob: Job; setSelectedId: (id: number) => void }) {
  return (
    <div style={styles.jobList}>
      <div style={styles.jobListHeader}><span>Job</span><span>Status</span><span>Phase</span><span>Risk</span><span>Budget Priority</span><span>Health</span></div>
      {jobs.map((job) => (
        <button key={job.id} style={{ ...styles.jobRow, background: job.id === selectedJob.id ? "#ecfdf5" : "white" }} onClick={() => setSelectedId(job.id)}>
          <div><strong style={styles.jobName}>{job.name}</strong><span style={styles.jobSub}>{job.jobNumber} · {job.customer} · {job.location}</span></div>
          <Badge tone="blue">{job.status}</Badge>
          <strong style={styles.tableText}>{job.phase}</strong>
          <Badge tone={job.risk === "Green" ? "green" : job.risk === "Red" ? "red" : "yellow"}>{job.risk}</Badge>
          <Badge tone={budgetPriorityTone(job)}>{budgetPriorityLabel(job)}</Badge>
          <MiniHealth job={job} />
        </button>
      ))}
    </div>
  );
}

function DetailGrid({ job }: { job: Job }) {
  return <div style={styles.detailGrid}><Detail label="Status" value={job.status} /><Detail label="Phase" value={job.phase} /><Detail label="Risk" value={job.risk} /><Detail label="Target Finish" value={job.finishDate} /></div>;
}

function StatusSnapshot({ job }: { job: Job }) {
  const taskDone = job.projectTasks.filter((task) => task.done).length;
  return (
    <div style={styles.statusPanel}>
      <div style={styles.compactCardHeader}><h3 style={styles.sectionTitle}>Project Status</h3><span style={styles.smallPill}>{taskDone}/{job.projectTasks.length} PM tasks</span></div>
      <div style={styles.statusGrid}>
        <StatusItem label="Current Phase" value={job.phase} />
        <StatusItem label="Overall Progress" value={`${job.progress}%`} />
        <StatusItem label="Risk" value={job.risk} />
        <StatusItem label="Target Finish" value={job.finishDate} />
        {job.inspections.map((inspection) => <StatusItem key={inspection.id} label={inspection.title} value={inspection.status} />)}
      </div>
    </div>
  );
}

function HealthBars({ job, isAdmin }: { job: Job; isAdmin: boolean }) {
  const labourUsed = pct(job.labourCostToDate || 0, job.labourBudget || 0);
  const budgetUsed = pct(job.costToDate, job.budget);
  return (
    <div style={styles.healthPanel}>
      <h3 style={styles.sectionTitle}>Job Health</h3>
      <Bar label="Job Progress" value={job.progress} helper={`${job.progress}%`} />
      <Bar label="Labour Used" value={labourUsed} helper={isAdmin ? `${currency(job.labourCostToDate || 0)} / ${currency(job.labourBudget || 0)} (${labourUsed}%)` : `${labourUsed}%`} warn={labourUsed > 85} />
      <Bar label="Labour Left" value={100 - labourUsed} helper={isAdmin ? `${currency(Math.max(0, (job.labourBudget || 0) - (job.labourCostToDate || 0)))} left (${100 - labourUsed}%)` : `${100 - labourUsed}%`} blue />
      <Bar label="Budget Health" value={100 - budgetUsed} helper={isAdmin ? `${currency(Math.max(0, job.budget - job.costToDate))} remaining (${100 - budgetUsed}%)` : `${100 - budgetUsed}%`} blue warn={budgetUsed > 85} />
    </div>
  );
}

function NextAction({ job, updateJob }: { job: Job; updateJob: (update: Partial<Job>) => void }) {
  return <div style={styles.nextActionBox}><div style={styles.nextActionTop}><h3 style={styles.sectionTitle}>Next Action</h3><span>Owner: {job.actionOwner} · Due: {job.actionDue}</span></div><textarea style={styles.nextActionInput} value={job.nextAction} onChange={(event) => updateJob({ nextAction: event.target.value })} /></div>;
}

function AdminJobEdit({ job, updateJob }: { job: Job; updateJob: (update: Partial<Job>) => void }) {
  return (
    <div style={styles.editBox}>
      <h3 style={styles.sectionTitle}>Admin Labour / Budget Inputs</h3>
      <div style={styles.editGrid}>
        <Field label="Status"><select style={styles.input} value={job.status} onChange={(event) => updateJob({ status: event.target.value as JobStatus })}><option>Planning</option><option>Active</option><option>On Hold</option><option>Complete</option></select></Field>
        <Field label="Progress %"><input style={styles.input} type="number" value={job.progress} onChange={(event) => updateJob({ progress: Number(event.target.value) })} /></Field>
        <Field label="Allowed Hours"><input style={styles.input} type="number" value={job.allowedHours} onChange={(event) => updateJob({ allowedHours: Number(event.target.value) })} /></Field>
        <Field label="Used Hours"><input style={styles.input} type="number" value={job.usedHours} onChange={(event) => updateJob({ usedHours: Number(event.target.value) })} /></Field>
        <Field label="Budget"><input style={styles.input} type="number" value={job.budget} onChange={(event) => updateJob({ budget: Number(event.target.value) })} /></Field>
        <Field label="Cost To Date"><input style={styles.input} type="number" value={job.costToDate} onChange={(event) => updateJob({ costToDate: Number(event.target.value) })} /></Field>
      </div>
    </div>
  );
}

function SmallJobsPage({ isAdmin, smallJobs, setSmallJobs }: { isAdmin: boolean; smallJobs: SmallJob[]; setSmallJobs: (jobs: SmallJob[]) => void }) {
  const today = new Date().toISOString().slice(0, 10);
  const emptyForm = { title: "", location: "", estimatedDays: 1, estimatedCrew: 1, priority: "Medium" as const, scope: "", documents: [] as JobDocumentUpload[], status: "Planning" as const, enteredDate: today, scheduled: false };
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState<number | null>(null);
  const [selectedSmallJobId, setSelectedSmallJobId] = useState<number | null>(smallJobs[0]?.id || null);
  const [form, setForm] = useState(emptyForm);
  const [previewDoc, setPreviewDoc] = useState<JobDocument | null>(null);

  if (!isAdmin) {
    return <PageTitle title="Jobs To Do" subtitle="Only Admin/PM can access this section." />;
  }

  const selectedSmallJob = smallJobs.find((job) => job.id === selectedSmallJobId) || smallJobs[0];

  function daysSinceEntered(date: string) {
    const start = new Date(`${date}T00:00:00`);
    const now = new Date(`${today}T00:00:00`);
    return Math.max(0, Math.floor((now.getTime() - start.getTime()) / 86400000));
  }

  function openAdd() {
    setEditingId(null);
    setForm(emptyForm);
    setShowForm(true);
  }

  function openEdit(job: SmallJob) {
    setEditingId(job.id);
    setForm({
      title: job.title,
      location: job.location,
      estimatedDays: job.estimatedDays,
      estimatedCrew: job.estimatedCrew,
      priority: job.priority,
      scope: job.scope,
      documents: job.documents,
      status: job.status,
      enteredDate: job.enteredDate || today,
      scheduled: job.scheduled || false,
    });
    setShowForm(true);
  }

  function saveSmallJob() {
    if (!form.title.trim()) return;

    const payload: SmallJob = {
      id: editingId || Date.now(),
      title: form.title.trim(),
      location: form.location || "TBD",
      estimatedDays: form.estimatedDays,
      estimatedCrew: form.estimatedCrew,
      priority: form.priority,
      scope: form.scope,
      documents: form.documents,
      status: form.status,
      enteredDate: form.enteredDate || today,
      scheduled: form.scheduled,
    };

    if (editingId) {
      setSmallJobs(smallJobs.map((job) => job.id === editingId ? payload : job));
    } else {
      setSmallJobs([payload, ...smallJobs]);
      setSelectedSmallJobId(payload.id);
    }

    setShowForm(false);
    setEditingId(null);
    setForm(emptyForm);
  }

  function deleteSmallJob(id: number) {
    const nextJobs = smallJobs.filter((job) => job.id !== id);
    setSmallJobs(nextJobs);
    if (selectedSmallJobId === id) setSelectedSmallJobId(nextJobs[0]?.id || null);
  }

  function toggleScheduled(id: number) {
    setSmallJobs(smallJobs.map((job) => job.id === id ? { ...job, scheduled: !job.scheduled } : job));
  }

  return (
    <div style={styles.cleanStack}>
      <Card>
        <div style={styles.compactCardHeader}>
          <div>
            <h2 style={styles.cardTitle}>Jobs To Do</h2>
            <p style={styles.muted}>Compact list for small jobs. Click a job name to view full details.</p>
          </div>
          <button style={styles.primary} onClick={openAdd}><Plus size={16} /> Add Small Job</button>
        </div>
      </Card>

      {showForm && (
        <Card>
          <div style={styles.compactCardHeader}>
            <h3 style={styles.sideTitle}>{editingId ? "Edit Job To Do" : "Add Job To Do"}</h3>
            <button style={styles.closeMiniButton} onClick={() => setShowForm(false)}>×</button>
          </div>

          <div style={styles.potentialJobFormGrid}>
            <Field label="Job / Task Name"><input style={styles.input} value={form.title} onChange={(event) => setForm({ ...form, title: event.target.value })} /></Field>
            <Field label="Location"><input style={styles.input} value={form.location} onChange={(event) => setForm({ ...form, location: event.target.value })} /></Field>
            <Field label="Estimated Days"><input style={styles.input} type="number" value={form.estimatedDays} onChange={(event) => setForm({ ...form, estimatedDays: Number(event.target.value) })} /></Field>
            <Field label="Estimated Crew"><input style={styles.input} type="number" value={form.estimatedCrew} onChange={(event) => setForm({ ...form, estimatedCrew: Number(event.target.value) })} /></Field>
            <Field label="Priority"><select style={styles.input} value={form.priority} onChange={(event) => setForm({ ...form, priority: event.target.value as SmallJob["priority"] })}><option>Low</option><option>Medium</option><option>High</option><option>Urgent</option></select></Field>
            <Field label="Status"><select style={styles.input} value={form.status} onChange={(event) => setForm({ ...form, status: event.target.value as SmallJob["status"] })}><option>Not Started</option><option>Planning</option><option>In Progress</option><option>Completed</option></select></Field>
            <Field label="Date Entered"><input style={styles.input} type="date" value={form.enteredDate} onChange={(event) => setForm({ ...form, enteredDate: event.target.value })} /></Field>
            <label style={styles.scheduledCheck}><input type="checkbox" checked={form.scheduled} onChange={(event) => setForm({ ...form, scheduled: event.target.checked })} /> Scheduled</label>
          </div>

          <Field label="Scope / Work Required"><textarea style={styles.dailyReportInput} value={form.scope} onChange={(event) => setForm({ ...form, scope: event.target.value })} placeholder="What work needs to be completed?" /></Field>

          <div style={styles.addInlineBox}>
            <h3 style={styles.sideTitle}>Documents / Drawings</h3>
            <label style={styles.dropZone}>
              <strong>Attach drawings or documents</strong>
              <span>PDFs, photos, screenshots, sketches, etc.</span>
              <input style={styles.hiddenFileInput} type="file" multiple onChange={(event) => {
                const files = Array.from(event.target.files || []);
                setForm({
                  ...form,
                  documents: [
                    ...form.documents,
                    ...files.map((file) => ({
                      fileName: file.name,
                      type: file.type.startsWith("image/") || file.name.toLowerCase().match(/\.(jpg|jpeg|png|gif|webp)$/) ? "Photo" as const : "Document" as const,
                      url: URL.createObjectURL(file),
                    })),
                  ],
                });
              }} />
            </label>
            <div style={styles.pictureList}>{form.documents.map((doc) => <button key={doc.fileName} style={styles.picturePreviewButton} onClick={() => setPreviewDoc({ id: Date.now(), jobId: 0, fileName: doc.fileName, uploadedBy: "Jobs To Do", uploadedDate: today, type: doc.type, url: doc.url })}>{doc.type === "Photo" ? "📷" : "📄"} {doc.fileName}</button>)}</div>
          </div>

          <div style={styles.modalActions}><button style={styles.secondary} onClick={() => setShowForm(false)}>Cancel</button><button style={styles.primary} onClick={saveSmallJob}>{editingId ? "Save Changes" : "Save Job"}</button></div>
        </Card>
      )}

      <div style={styles.smallJobsSplitLayout}>
        <Card>
          <h3 style={styles.sideTitle}>Small Jobs List</h3>
          <div style={styles.smallJobsList}>
            {smallJobs.map((job) => {
              const daysOpen = daysSinceEntered(job.enteredDate);
              const needsAttention = !job.scheduled && job.status !== "Completed" && daysOpen > 15;
              return (
                <button key={job.id} style={selectedSmallJob?.id === job.id ? styles.smallJobRowActive : styles.smallJobRow} onClick={() => setSelectedSmallJobId(job.id)}>
                  <div style={styles.smallJobRowTitle}><strong>{job.title}</strong>{needsAttention && <span style={styles.smallJobAlertDot}>!</span>}</div>
                  <span>{job.location} · {job.status}</span>
                  <div style={styles.smallJobRowMeta}>
                    <Badge tone={job.priority === "Urgent" || job.priority === "High" ? "red" : job.priority === "Medium" ? "yellow" : "green"}>{job.priority}</Badge>
                    <span>{job.estimatedDays} day(s)</span>
                    <span>{job.estimatedCrew} guy(s)</span>
                    <span>{job.scheduled ? "Scheduled" : `${daysOpen} days open`}</span>
                  </div>
                </button>
              );
            })}
          </div>
        </Card>

        <Card>
          {selectedSmallJob ? (() => {
            const daysOpen = daysSinceEntered(selectedSmallJob.enteredDate);
            const needsAttention = !selectedSmallJob.scheduled && selectedSmallJob.status !== "Completed" && daysOpen > 15;
            return (
              <div style={styles.cleanStack}>
                <div style={styles.compactCardHeader}>
                  <div>
                    <h3 style={styles.cardTitle}>{selectedSmallJob.title}</h3>
                    <p style={styles.muted}>{selectedSmallJob.location}</p>
                  </div>
                  <Badge tone={selectedSmallJob.priority === "Urgent" || selectedSmallJob.priority === "High" ? "red" : selectedSmallJob.priority === "Medium" ? "yellow" : "green"}>{selectedSmallJob.priority}</Badge>
                </div>

                {needsAttention && <div style={styles.alertNotice}>⚠ Entered {daysOpen} days ago with no scheduled action.</div>}

                <div style={styles.potentialJobStats}>
                  <Detail label="Days" value={`${selectedSmallJob.estimatedDays}`} />
                  <Detail label="Crew" value={`${selectedSmallJob.estimatedCrew}`} />
                  <Detail label="Status" value={selectedSmallJob.status} />
                  <Detail label="Entered" value={selectedSmallJob.enteredDate} />
                  <Detail label="Scheduled" value={selectedSmallJob.scheduled ? "Yes" : "No"} />
                  <Detail label="Open" value={`${daysOpen} days`} />
                </div>

                <div style={styles.scopeBox}>{selectedSmallJob.scope || "No description entered yet."}</div>

                <div style={styles.pictureList}>{selectedSmallJob.documents.map((doc) => <button key={doc.fileName} style={styles.picturePreviewButton} onClick={() => setPreviewDoc({ id: Date.now(), jobId: 0, fileName: doc.fileName, uploadedBy: "Jobs To Do", uploadedDate: today, type: doc.type, url: doc.url })}>{doc.type === "Photo" ? "📷" : "📄"} {doc.fileName}</button>)}</div>

                <div style={styles.smallJobActions}>
                  <label style={styles.scheduledCheck}><input type="checkbox" checked={selectedSmallJob.scheduled} onChange={() => toggleScheduled(selectedSmallJob.id)} /> Scheduled</label>
                  <button style={styles.secondary} onClick={() => openEdit(selectedSmallJob)}>Edit Job</button>
                  <button style={styles.danger} onClick={() => deleteSmallJob(selectedSmallJob.id)}><Trash2 size={15} /> Delete</button>
                </div>
              </div>
            );
          })() : <p style={styles.muted}>Select a job to view details.</p>}
        </Card>
      </div>

      {previewDoc && <DocumentPreviewModal document={previewDoc} onClose={() => setPreviewDoc(null)} />}
    </div>
  );
}

function CalendarView({ jobs, potentialJobs, isAdmin }: { jobs: Job[]; potentialJobs: PotentialJob[]; isAdmin: boolean }) {
  const [monthAnchor, setMonthAnchor] = useState(new Date());
  const year = monthAnchor.getFullYear();
  const month = monthAnchor.getMonth();
  const firstDay = new Date(year, month, 1);
  const startOffset = firstDay.getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const cells = Array.from({ length: 42 }, (_, index) => {
    const dayNumber = index - startOffset + 1;
    const date = new Date(year, month, dayNumber);
    const dateKey = date.toISOString().slice(0, 10);
    const inMonth = dayNumber >= 1 && dayNumber <= daysInMonth;
    const awardedOnDay = jobs.filter((job) => job.startDate <= dateKey && job.finishDate >= dateKey);
    const potentialOnDay = isAdmin ? potentialJobs.filter((job) => job.startDate <= dateKey && job.finishDate >= dateKey) : [];
    return { dayNumber, dateKey, inMonth, awardedOnDay, potentialOnDay };
  });

  function shiftMonth(amount: number) {
    setMonthAnchor(new Date(year, month + amount, 1));
  }

  return (
    <div style={styles.cleanStack}>
      <Card>
        <div style={styles.compactCardHeader}>
          <div>
            <h2 style={styles.cardTitle}>Job Calendar</h2>
            <p style={styles.muted}>Awarded jobs show solid colours. Potential jobs show lighter transparent cards.</p>
          </div>
          <div style={styles.calendarControls}>
            <button style={styles.secondary} onClick={() => shiftMonth(-1)}>← Previous</button>
            <strong>{monthAnchor.toLocaleString("default", { month: "long" })} {year}</strong>
            <button style={styles.secondary} onClick={() => shiftMonth(1)}>Next →</button>
          </div>
        </div>
        <div style={styles.calendarLegend}>
          <span style={styles.awardedLegend}>Awarded Job</span>
          {isAdmin && <span style={styles.potentialLegend}>Potential Job</span>}
        </div>
      </Card>

      <div style={styles.calendarGrid}>
        {["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].map((day) => <div key={day} style={styles.calendarHeaderCell}>{day}</div>)}
        {cells.map((cell) => (
          <div key={cell.dateKey} style={cell.inMonth ? styles.calendarCell : styles.calendarCellMuted}>
            <strong style={styles.calendarDay}>{cell.inMonth ? cell.dayNumber : ""}</strong>
            <div style={styles.calendarItems}>
              {cell.awardedOnDay.map((job, index) => <div key={`${job.id}-${cell.dateKey}`} style={{ ...styles.calendarJobItem, borderLeftColor: index % 3 === 0 ? "#047857" : index % 3 === 1 ? "#2563eb" : "#7c3aed" }}>{job.name}</div>)}
              {cell.potentialOnDay.map((job) => <div key={`p-${job.id}-${cell.dateKey}`} style={styles.calendarPotentialItem}>{job.name}</div>)}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function PotentialJobsPage({ isAdmin, currentUser, potentialJobs, setPotentialJobs }: { isAdmin: boolean; currentUser: User; potentialJobs: PotentialJob[]; setPotentialJobs: (jobs: PotentialJob[]) => void }) {
  const today = new Date().toISOString().slice(0, 10);
  const emptyForm = { name: "", customer: "", location: "", startDate: today, finishDate: today, estimatedHours: 0, assumedValue: 0, probability: "Medium", documents: [] as JobDocumentUpload[] };
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState<number | null>(null);
  const [form, setForm] = useState(emptyForm);
  const [previewDoc, setPreviewDoc] = useState<JobDocument | null>(null);

  if (!isAdmin) {
    return <PageTitle title="Potential Jobs" subtitle="Only Admin/PM can view potential jobs." />;
  }

  function startAdd() {
    setForm(emptyForm);
    setEditingId(null);
    setShowForm(true);
  }

  function startEdit(job: PotentialJob) {
    setForm({
      name: job.name,
      customer: job.customer,
      location: job.location,
      startDate: job.startDate,
      finishDate: job.finishDate,
      estimatedHours: job.estimatedHours,
      assumedValue: job.assumedValue || 0,
      probability: job.probability,
      documents: job.documents || [],
    });
    setEditingId(job.id);
    setShowForm(true);
  }

  function savePotentialJob() {
    if (!form.name.trim()) return;
    const todayStamp = new Date().toISOString().slice(0, 10);

    if (editingId) {
      setPotentialJobs(potentialJobs.map((job) => {
        if (job.id !== editingId) return job;

        const nextJob: PotentialJob = {
          ...job,
          name: form.name.trim(),
          customer: form.customer || "TBD",
          location: form.location || "TBD",
          probability: form.probability,
          startDate: form.startDate,
          finishDate: form.finishDate,
          estimatedHours: form.estimatedHours,
          assumedValue: form.assumedValue,
          document: form.documents[0]?.fileName || job.document,
          documents: form.documents,
          activityLog: job.activityLog || [],
        };

        const fieldsToCheck: { field: string; previousValue: string; newValue: string }[] = [
          { field: "Job Name", previousValue: job.name, newValue: nextJob.name },
          { field: "Customer / GC", previousValue: job.customer, newValue: nextJob.customer },
          { field: "Location", previousValue: job.location, newValue: nextJob.location },
          { field: "Probability", previousValue: job.probability, newValue: nextJob.probability },
          { field: "Assumed Start", previousValue: job.startDate, newValue: nextJob.startDate },
          { field: "Assumed Finish", previousValue: job.finishDate, newValue: nextJob.finishDate },
          { field: "Assumed Hours", previousValue: String(job.estimatedHours), newValue: String(nextJob.estimatedHours) },
          { field: "Assumed Value", previousValue: currency(job.assumedValue || 0), newValue: currency(nextJob.assumedValue || 0) },
          { field: "Documents", previousValue: `${(job.documents || []).length} attached`, newValue: `${(nextJob.documents || []).length} attached` },
        ];

        const changes = fieldsToCheck
          .filter((item) => item.previousValue !== item.newValue)
          .map((item, index) => ({
            id: Date.now() + index,
            date: todayStamp,
            user: currentUser.name,
            field: item.field,
            previousValue: item.previousValue || "—",
            newValue: item.newValue || "—",
          }));

        return { ...nextJob, activityLog: [...changes, ...(job.activityLog || [])] };
      }));
    } else {
      const originalSummary = [
        `Name: ${form.name.trim()}`,
        `Customer: ${form.customer || "TBD"}`,
        `Location: ${form.location || "TBD"}`,
        `Probability: ${form.probability}`,
        `Dates: ${form.startDate} to ${form.finishDate}`,
        `Hours: ${form.estimatedHours}`,
        `Value: ${currency(form.assumedValue || 0)}`,
      ].join(" | ");

      setPotentialJobs([
        {
          id: Date.now(),
          name: form.name.trim(),
          customer: form.customer || "TBD",
          location: form.location || "TBD",
          probability: form.probability,
          startDate: form.startDate,
          finishDate: form.finishDate,
          crewNeeded: 0,
          estimatedHours: form.estimatedHours,
          assumedValue: form.assumedValue,
          crewIds: [],
          document: form.documents[0]?.fileName || "",
          documents: form.documents,
          activityLog: [{ id: Date.now() + 1, date: todayStamp, user: currentUser.name, field: "Original Entry", previousValue: "—", newValue: originalSummary }],
        },
        ...potentialJobs,
      ]);
    }
    setShowForm(false);
    setEditingId(null);
    setForm(emptyForm);
  }

  function deletePotentialJob(id: number) {
    setPotentialJobs(potentialJobs.filter((job) => job.id !== id));
  }

  return (
    <div style={styles.cleanStack}>
      <Card>
        <div style={styles.compactCardHeader}>
          <div>
            <h2 style={styles.cardTitle}>Potential Jobs</h2>
            <p style={styles.muted}>Track possible jobs, assumed dates, assumed hours, assumed value, and attached drawings/documents.</p>
          </div>
          <button style={styles.primary} onClick={startAdd}><Plus size={16} /> Add Potential Job</button>
        </div>
      </Card>

      {showForm && (
        <Card>
          <div style={styles.compactCardHeader}>
            <h3 style={styles.sideTitle}>{editingId ? "Edit Potential Job" : "Add Potential Job"}</h3>
            <button style={styles.closeMiniButton} onClick={() => setShowForm(false)}>×</button>
          </div>
          <div style={styles.potentialJobFormGrid}>
            <Field label="Job Name"><input style={styles.input} value={form.name} onChange={(event) => setForm({ ...form, name: event.target.value })} /></Field>
            <Field label="Customer / GC"><input style={styles.input} value={form.customer} onChange={(event) => setForm({ ...form, customer: event.target.value })} /></Field>
            <Field label="Location"><input style={styles.input} value={form.location} onChange={(event) => setForm({ ...form, location: event.target.value })} /></Field>
            <Field label="Assumed Start"><input style={styles.input} type="date" value={form.startDate} onChange={(event) => setForm({ ...form, startDate: event.target.value })} /></Field>
            <Field label="Assumed Finish"><input style={styles.input} type="date" value={form.finishDate} onChange={(event) => setForm({ ...form, finishDate: event.target.value })} /></Field>
            <Field label="Assumed Hours"><input style={styles.input} type="number" value={form.estimatedHours} onChange={(event) => setForm({ ...form, estimatedHours: Number(event.target.value) })} /></Field>
            <Field label="Assumed Value"><input style={styles.input} type="number" value={form.assumedValue} onChange={(event) => setForm({ ...form, assumedValue: Number(event.target.value) })} /></Field>
            <Field label="Probability"><select style={styles.input} value={form.probability} onChange={(event) => setForm({ ...form, probability: event.target.value })}><option>Low</option><option>Medium</option><option>High</option><option>Likely Awarded</option></select></Field>
          </div>
          <div style={styles.addInlineBox}>
            <h3 style={styles.sideTitle}>Drawings / Documents</h3>
            <label style={styles.dropZone}>
              <strong>Drag/drop or click to attach drawings/documents</strong>
              <span>PDF, Excel, images, or tender documents</span>
              <input style={styles.hiddenFileInput} type="file" multiple onChange={(event) => {
                const files = Array.from(event.target.files || []);
                setForm({
                  ...form,
                  documents: [
                    ...form.documents,
                    ...files.map((file) => ({
                      fileName: file.name,
                      type: file.type.startsWith("image/") || file.name.toLowerCase().match(/\.(jpg|jpeg|png|gif|webp)$/) ? "Photo" as const : "Document" as const,
                      url: URL.createObjectURL(file),
                    })),
                  ],
                });
              }} />
            </label>
            <div style={styles.pictureList}>{form.documents.map((doc) => <button key={doc.fileName} style={styles.picturePreviewButton} onClick={() => setPreviewDoc({ id: Date.now(), jobId: 0, fileName: doc.fileName, uploadedBy: "Potential Job", uploadedDate: today, type: doc.type, url: doc.url })}>{doc.type === "Photo" ? "📷" : "📄"} {doc.fileName}</button>)}</div>
          </div>
          <div style={styles.modalActions}><button style={styles.secondary} onClick={() => setShowForm(false)}>Cancel</button><button style={styles.primary} onClick={savePotentialJob}>{editingId ? "Save Changes" : "Save Potential Job"}</button></div>
        </Card>
      )}

      <div style={styles.potentialJobGrid}>
        {potentialJobs.map((job) => (
          <div style={styles.potentialJobCard} key={job.id}>
            <div style={styles.compactCardHeader}>
              <div><strong>{job.name}</strong><span>{job.customer} · {job.location}</span></div>
              <Badge tone={job.probability === "High" || job.probability === "Likely Awarded" ? "green" : job.probability === "Low" ? "red" : "yellow"}>{job.probability}</Badge>
            </div>
            <div style={{ ...styles.potentialJobStats, fontSize: 11 }}>
              <Detail label="Start" value={job.startDate} />
              <Detail label="Finish" value={job.finishDate} />
              <Detail label="Hours" value={`${job.estimatedHours}`} />
              <Detail label="Value" value={currency(job.assumedValue || 0)} />
            </div>
            <div style={styles.pictureList}>{(job.documents || []).map((doc) => <button key={doc.fileName} style={styles.picturePreviewButton} onClick={() => setPreviewDoc({ id: Date.now(), jobId: 0, fileName: doc.fileName, uploadedBy: "Potential Job", uploadedDate: job.startDate, type: doc.type, url: doc.url })}>{doc.type === "Photo" ? "📷" : "📄"} {doc.fileName}</button>)}</div>
            <div style={styles.potentialActivityBox}>
              <strong>Activity Log</strong>
              {(job.activityLog || []).length === 0 ? <span>No changes recorded yet.</span> : (job.activityLog || []).slice(0, 6).map((entry) => (
                <div style={styles.potentialActivityItem} key={entry.id}>
                  <span>{entry.date} · {entry.user}</span>
                  <b>{entry.field}</b>
                  <small>{entry.previousValue} → {entry.newValue}</small>
                </div>
              ))}
            </div>
            <div style={styles.smallJobActions}>
              <button style={styles.secondary} onClick={() => startEdit(job)}>Edit Potential Job</button>
              <button style={styles.danger} onClick={() => deletePotentialJob(job.id)}><Trash2 size={15} /> Delete</button>
            </div>
          </div>
        ))}
      </div>
      {previewDoc && <DocumentPreviewModal document={previewDoc} onClose={() => setPreviewDoc(null)} />}
    </div>
  );
}

function PlanningView({ isAdmin, jobs, visibleJobs, potentialJobs, setPotentialJobs, users }: { isAdmin: boolean; jobs: Job[]; visibleJobs: Job[]; potentialJobs: PotentialJob[]; setPotentialJobs: (jobs: PotentialJob[]) => void; users: User[] }) {
  return (
    <div style={styles.cleanStack}>
      <PageTitle title="Planning" subtitle={isAdmin ? "Awarded jobs, potential jobs, crew scheduling, long-lead tracking, and site readiness." : "Assigned job planning, crew schedule, long-lead status, and site readiness."} />
      <div style={styles.moduleGrid}>
        <ModuleCard title="Awarded Job Planning" icon={<BriefcaseBusiness />} items={visibleJobs.map((job) => `${job.name} · ${job.phase} · ${job.progress}%`)} />
        {isAdmin && <PotentialJobsCard users={users} jobs={jobs} potentialJobs={potentialJobs} setPotentialJobs={setPotentialJobs} />}
        <ModuleCard title="Crew Scheduling" icon={<Users />} items={["Weekly manpower board", "Foreman/site assignment", "Apprentice pairing", "Travel days and rotations"]} />
        <ModuleCard title="Long Lead Tracking" icon={<PackageCheck />} items={["Switchgear / panels", "Lighting package", "Generator / transfer switch", "Transformers / cable tray"]} />
        <ModuleCard title="Site Readiness" icon={<ShieldCheck />} items={["Temp power ready", "Permits pulled", "Lift/rentals arranged", "Material staged for mobilization"]} />
      </div>
    </div>
  );
}

function PotentialJobsCard({ users, jobs, potentialJobs, setPotentialJobs }: { users: User[]; jobs: Job[]; potentialJobs: PotentialJob[]; setPotentialJobs: (jobs: PotentialJob[]) => void }) {
  function update(id: number, updateValue: Partial<PotentialJob>) {
    setPotentialJobs(potentialJobs.map((job) => (job.id === id ? { ...job, ...updateValue } : job)));
  }
  return (
    <Card style={styles.wideModule}>
      <div style={styles.compactCardHeader}><h3 style={styles.sideTitle}>Potential Jobs</h3><span style={styles.smallPill}>Admin/PM</span></div>
      {potentialJobs.map((job) => {
        const overlaps = jobs.flatMap((awarded) => job.crewIds.filter((id) => awarded.crewIds.includes(id)).map((id) => `${userById(users, id)?.name} overlaps with ${awarded.name}`));
        return <div key={job.id} style={styles.potentialItem}><input style={styles.input} value={job.name} onChange={(e) => update(job.id, { name: e.target.value })} /><div style={styles.potentialSmallGrid}><Field label="Probability"><select style={styles.input} value={job.probability} onChange={(e) => update(job.id, { probability: e.target.value })}><option>Low</option><option>Medium</option><option>High</option><option>Likely Awarded</option></select></Field><Field label="Crew"><input style={styles.input} type="number" value={job.crewNeeded} onChange={(e) => update(job.id, { crewNeeded: Number(e.target.value) })} /></Field><Field label="Hours"><input style={styles.input} type="number" value={job.estimatedHours} onChange={(e) => update(job.id, { estimatedHours: Number(e.target.value) })} /></Field></div>{overlaps.length > 0 && <div style={styles.warning}>{overlaps.slice(0, 2).map((line) => <div key={line}>⚠ {line}</div>)}</div>}</div>;
      })}
    </Card>
  );
}

function ReportsView({ isAdmin, jobs }: { isAdmin: boolean; jobs: Job[] }) {
  return <div style={styles.cleanStack}><PageTitle title="Reports" subtitle="Daily reports, site reports, inspection reports, and export packages." /><div style={styles.moduleGrid}><ModuleCard title="Daily Reports" icon={<FileText />} items={["Foreman daily notes", "Labour installed", "Delays / issues", "Site photos"]} /><ModuleCard title="Site Reports" icon={<FileText />} items={["Progress reports", "Consultant/RFI issues", "Client-ready summaries", "Deficiency notes"]} /><ModuleCard title="Inspection Reports" icon={<ClipboardCheck />} items={jobs.flatMap((job) => job.inspections.map((inspection) => `${job.name}: ${inspection.title} · ${inspection.status}`)).slice(0, 6)} />{isAdmin && <ModuleCard title="Financial Reports" icon={<BarChart3 />} items={["Labour cost", "Profitability", "Change order totals", "Cost to complete"]} />}<ModuleCard title="Export Center" icon={<FileText />} items={["PDF report package", "Excel export", "Inspection summary", "O&M / closeout package"]} /></div></div>;
}

function CrewTasksView({ jobs, currentUser, toggleCrewTaskForJob, addCrewTaskForJob, submitDailyReport, dailyReports }: { jobs: Job[]; currentUser: User; toggleCrewTaskForJob: (jobId: number, taskId: number) => void; addCrewTaskForJob: (jobId: number, title: string) => void; submitDailyReport: (jobId: number, note: string, pictures?: string[], pictureUploads?: JobDocumentUpload[]) => void; dailyReports: DailyReport[] }) {
  const [taskTextByJob, setTaskTextByJob] = useState<Record<number, string>>({});
  const [reportTextByJob, setReportTextByJob] = useState<Record<number, string>>({});
  const earnsDailyReportPoints = currentUser.role === "Crew" || currentUser.role === "Foreman";

  return (
    <div style={styles.cleanStack}>
      <PageTitle
        title="Crew Tasks"
        subtitle={earnsDailyReportPoints ? "You only see crew tasks for jobs you are assigned to. Submit a daily report to earn 5 crew points." : "Admin/PM can view crew tasks for all jobs."}
      />

      {earnsDailyReportPoints && (
        <Card>
          <div style={styles.allowanceTop}>
            <div>
              <h2 style={styles.cardTitle}>Daily Report Points</h2>
              <p style={styles.muted}>Every daily report submitted by crew or foreman automatically adds 5 points to that user.</p>
            </div>
            <div style={styles.allowanceBadge}>+5 pts/report</div>
          </div>
        </Card>
      )}

      <div style={styles.assignedTaskGrid}>
        {jobs.map((job) => {
          const taskText = taskTextByJob[job.id] || "";
          const reportText = reportTextByJob[job.id] || "";
          const myReports = dailyReports.filter((report) => report.jobId === job.id && report.userId === currentUser.id);
          return (
            <Card key={job.id}>
              <div style={styles.compactCardHeader}>
                <div>
                  <h3 style={styles.sideTitle}>{job.name}</h3>
                  <p style={styles.cardHelp}>{job.phase} · {job.progress}% complete</p>
                </div>
                <span style={styles.smallPill}>{job.crewTasks.filter((task) => !task.done).length} open</span>
              </div>

              <div style={styles.pmRequestForm}>
                <input style={styles.input} value={taskText} onChange={(event) => setTaskTextByJob((all) => ({ ...all, [job.id]: event.target.value }))} placeholder="Add a crew task for this job..." />
                <button style={styles.smallButton} onClick={() => { addCrewTaskForJob(job.id, taskText); setTaskTextByJob((all) => ({ ...all, [job.id]: "" })); }}><Plus size={16} /> Add</button>
              </div>

              <TaskCard title="Crew Task Summary" tasks={job.crewTasks} onToggle={(id) => toggleCrewTaskForJob(job.id, id)} canEdit />

              <div style={styles.dailyReportBox}>
                <div style={styles.compactCardHeader}>
                  <h3 style={styles.sideTitle}>Daily Report</h3>
                  {earnsDailyReportPoints && <span style={styles.smallPill}>+5 pts</span>}
                </div>
                <textarea style={styles.dailyReportInput} value={reportText} onChange={(event) => setReportTextByJob((all) => ({ ...all, [job.id]: event.target.value }))} placeholder="Enter daily report notes, delays, work completed, materials needed, or issues..." />
                <button style={styles.primary} onClick={() => { submitDailyReport(job.id, reportText); setReportTextByJob((all) => ({ ...all, [job.id]: "" })); }}>Submit Daily Report</button>
                <div style={styles.reportHistoryList}>
                  {myReports.slice(0, 3).map((report) => (
                    <div style={styles.reportHistoryItem} key={report.id}>
                      <strong>{report.date} · +{report.pointsAwarded} pts</strong>
                      <span>{report.note}</span>
                    </div>
                  ))}
                </div>
              </div>
            </Card>
          );
        })}
      </div>
    </div>
  );
}

function CrewPoints({ users, setUsers, jobs, points, currentUser, selectedJob, awardPoints, deductPoints, isAdmin, isForeman }: { users: User[]; setUsers: (users: User[]) => void; jobs: Job[]; points: PointHistory[]; currentUser: User; selectedJob: Job; awardPoints: (id: number, points: number, reason: string, jobName?: string) => void; deductPoints: (id: number, points: number, reason: string) => void; isAdmin: boolean; isForeman: boolean }) {
  const [selectedCrewId, setSelectedCrewId] = useState(selectedJob.crewIds.find((id) => userById(users, id)?.role === "Crew") || users.find((user) => user.role === "Crew")?.id || 5);
  const assignedJobsForSelectedCrew = jobs.filter((job) => job.crewIds.includes(selectedCrewId));
  const [selectedPointJobName, setSelectedPointJobName] = useState(assignedJobsForSelectedCrew[0]?.name || selectedJob.name);
  const [amount, setAmount] = useState(10);
  const [reason, setReason] = useState("");
  const [redeemUserId, setRedeemUserId] = useState(users.find((user) => user.role === "Crew" || user.role === "Foreman")?.id || selectedCrewId);
  const [redeemAmount, setRedeemAmount] = useState(10);
  const [redeemReason, setRedeemReason] = useState("");
  const awardedByForeman = points.filter((point) => point.awardedById === currentUser.id).reduce((sum, point) => sum + point.points, 0);
  const foremanRemaining = Math.max(0, currentUser.monthlyPointLimit - awardedByForeman);
  const awardableUsers = isAdmin ? users.filter((user) => user.role !== "Admin") : users.filter((user) => user.role === "Crew" && selectedJob.crewIds.includes(user.id));
  const visibleUsers = currentUser.role === "Crew" || isForeman ? users.filter((user) => user.id === currentUser.id) : isAdmin ? users : [];
  const pointJobOptions = assignedJobsForSelectedCrew.length > 0 ? assignedJobsForSelectedCrew : [selectedJob];

  function changeSelectedCrew(id: number) {
    setSelectedCrewId(id);
    const firstAssignedJob = jobs.find((job) => job.crewIds.includes(id));
    setSelectedPointJobName(firstAssignedJob?.name || selectedJob.name);
  }

  return (
    <div style={styles.cleanStack}>
      <PageTitle title="Crew Points" subtitle={currentUser.role === "Crew" ? "Crew members only see their own point total and history." : isForeman ? `Foremen can give points to assigned crew only. ${foremanRemaining} pts left to give.` : "Admin/PM can see totals, set foreman limits, and award points."} />
      {(isAdmin || isForeman) && (
        <Card>
          <h2 style={styles.cardTitle}>Award Points</h2>
          <div style={styles.pointsFormWithJob}>
            <Field label="Crew Member"><select style={styles.input} value={selectedCrewId} onChange={(event) => changeSelectedCrew(Number(event.target.value))}>{awardableUsers.map((user) => <option key={user.id} value={user.id}>{user.name}</option>)}</select></Field>
            <Field label="Job"><select style={styles.input} value={selectedPointJobName} onChange={(event) => setSelectedPointJobName(event.target.value)}>{pointJobOptions.map((job) => <option key={job.id} value={job.name}>{job.name}</option>)}</select></Field>
            <Field label="Points"><input style={styles.input} type="number" value={amount} max={isForeman ? foremanRemaining : undefined} onChange={(event) => setAmount(Number(event.target.value))} /></Field>
            <Field label="Reason"><input style={styles.input} value={reason} onChange={(event) => setReason(event.target.value)} placeholder="Reason for points" /></Field>
            <button style={styles.primary} onClick={() => { awardPoints(selectedCrewId, amount, reason, selectedPointJobName); setReason(""); }}>Save Points</button>
          </div>
        </Card>
      )}
      {isAdmin && (
        <Card>
          <h2 style={styles.cardTitle}>Redeem / Deduct Crew Points</h2>
          <p style={styles.cardHelp}>Use this when a crew member spends points on PJ's merch, tools, gift cards, lunches, or other rewards. This subtracts points and keeps a history line.</p>
          <div style={styles.pointsFormWithJob}>
            <Field label="Crew Member"><select style={styles.input} value={redeemUserId} onChange={(event) => setRedeemUserId(Number(event.target.value))}>{users.filter((user) => user.role === "Crew" || user.role === "Foreman").map((user) => <option key={user.id} value={user.id}>{user.name} — {user.points} pts</option>)}</select></Field>
            <Field label="Points To Use"><input style={styles.input} type="number" value={redeemAmount} onChange={(event) => setRedeemAmount(Number(event.target.value))} /></Field>
            <Field label="Purchase / Reason"><input style={styles.input} value={redeemReason} onChange={(event) => setRedeemReason(event.target.value)} placeholder="Example: Hoodie, lunch, gift card..." /></Field>
            <button style={styles.danger} onClick={() => { deductPoints(redeemUserId, redeemAmount, redeemReason); setRedeemReason(""); }}>Use Points</button>
          </div>
        </Card>
      )}
      {isAdmin && <Card><h2 style={styles.cardTitle}>Foreman Point Limits</h2><div style={styles.limitRows}>{users.filter((user) => user.role === "Foreman").map((user) => <div key={user.id} style={styles.limitRow}><div><strong>{user.name}</strong><span>Monthly point allowance admin allows</span></div><input style={styles.input} type="number" value={user.monthlyPointLimit} onChange={(event) => setUsers(users.map((existing) => existing.id === user.id ? { ...existing, monthlyPointLimit: Number(event.target.value) } : existing))} /></div>)}</div></Card>}
      {isForeman && <Card><div style={styles.allowanceTop}><h2 style={styles.cardTitle}>My Award Allowance</h2><div style={styles.allowanceBadge}>{foremanRemaining} pts left</div></div><Bar label="Allowance Used" value={pct(awardedByForeman, currentUser.monthlyPointLimit)} helper={`${awardedByForeman} of ${currentUser.monthlyPointLimit} pts used`} /><p style={styles.muted}>Foremen cannot see crew member total point balances.</p></Card>}
      <Card><h2 style={styles.cardTitle}>{currentUser.role === "Crew" || isForeman ? "My Points" : "Company Point Totals"}</h2><div style={styles.pointsGrid}>{visibleUsers.map((user) => { const history = points.filter((point) => point.userId === user.id); return <div key={user.id} style={styles.pointCard}><div style={styles.pointTop}><div><strong>{user.name}</strong><span>{user.role}</span></div><b>{user.points} pts</b></div><div style={styles.pointHistory}>{history.length ? history.map((point) => <div key={point.id}>+{point.points} · {point.reason}<small>{point.date} · {point.jobName}</small></div>) : <span>No points yet.</span>}</div></div>; })}</div></Card>
    </div>
  );
}

function InspectionsView({ jobs, selectedJob, setSelectedId, updateInspection, isAdmin }: { jobs: Job[]; selectedJob: Job; setSelectedId: (id: number) => void; updateInspection: (id: number, status: string) => void; isAdmin: boolean }) {
  return <div style={styles.cleanStack}><PageTitle title="Inspections" subtitle={isAdmin ? "Admin/PM can update inspection statuses." : "Crew and foremen can view inspection status only."} /><div style={styles.dashboardGrid}><Card><h3 style={styles.sideTitle}>Jobs</h3>{jobs.map((job) => <button key={job.id} style={{ ...styles.cleanInspection, width: "100%", textAlign: "left", cursor: "pointer", background: job.id === selectedJob.id ? "#ecfdf5" : "#f8fafc" }} onClick={() => setSelectedId(job.id)}><strong>{job.name}</strong><span>{job.inspections.length} inspections</span></button>)}</Card><InspectionCard inspections={selectedJob.inspections} canEdit={isAdmin} onChange={updateInspection} /></div></div>;
}

function TimeTrackingView({ isAdmin, jobs, currentUser }: { isAdmin: boolean; jobs: Job[]; currentUser: User }) {
  return <div style={styles.cleanStack}><PageTitle title="Time Tracking" subtitle="Clock time, PO/PCN coding, notes, and job hour imports." /><div style={styles.moduleGrid}><ModuleCard title="Clock Entry" icon={<Timer />} items={[`User: ${currentUser.name}`, "Select job", "Select PO / PCN", "Add notes and photos"]} /><ModuleCard title="Foreman Review" icon={<Users />} items={["Review crew hours", "Check daily notes", "Approve site photos", "Flag missing time"]} />{isAdmin && <ModuleCard title="Knowify Import" icon={<FileText />} items={["Import time report", "Prevent duplicate date ranges", "Map hours to job", "Review skipped hours"]} />}<ModuleCard title="Assigned Job Hours" icon={<BarChart3 />} items={jobs.map((job) => `${job.name}: ${isAdmin ? `${job.usedHours} hrs` : `${pct(job.usedHours, job.allowedHours)}% used`}`)} /></div></div>;
}

function LabourTrackingView({ isAdmin, jobs }: { isAdmin: boolean; jobs: Job[] }) {
  return <div style={styles.cleanStack}><PageTitle title="Labour Tracking" subtitle={isAdmin ? "Admin view with true labour and budget values." : "Crew/foreman view shows percentages only."} /><div style={styles.labourGrid}>{jobs.map((job) => <Card key={job.id}><h3 style={styles.sideTitle}>{job.name}</h3><HealthBars job={job} isAdmin={isAdmin} />{isAdmin && <div style={styles.detailGrid}><Detail label="Allowed Hours" value={`${job.allowedHours}`} /><Detail label="Used Hours" value={`${job.usedHours}`} /><Detail label="Budget" value={currency(job.budget)} /><Detail label="Cost To Date" value={currency(job.costToDate)} /></div>}</Card>)}</div></div>;
}

function AddJobModal({ users, onClose, onCreate }: { users: User[]; onClose: () => void; onCreate: (input: { name: string; customer: string; location: string; startDate: string; finishDate: string; crewIds: number[]; labourBudget: number; budget: number }) => void }) {
  const today = new Date().toISOString().slice(0, 10);
  const [name, setName] = useState("");
  const [customer, setCustomer] = useState("");
  const [location, setLocation] = useState("");
  const [startDate, setStartDate] = useState(today);
  const [finishDate, setFinishDate] = useState(today);
  const [labourBudget, setLabourBudget] = useState(0);
  const [budget, setBudget] = useState(0);
  const [crewIds, setCrewIds] = useState<number[]>([]);
  const assignableUsers = users.filter((user) => user.role === "Foreman" || user.role === "Crew");

  function toggleCrew(id: number) {
    setCrewIds((current) => current.includes(id) ? current.filter((crewId) => crewId !== id) : [...current, id]);
  }

  return (
    <div style={styles.modalBackdrop}>
      <div style={styles.settingsModal}>
        <div style={styles.cardHeader}>
          <div>
            <h2 style={styles.cardTitle}>Add New Job</h2>
            <p style={styles.muted}>Set up the job name, dates, assigned employees, and preload the standard PM checklist.</p>
          </div>
          <button style={styles.closeButton} onClick={onClose}>×</button>
        </div>

        <Card>
          <div style={styles.addJobGrid}>
            <Field label="Job Name"><input style={styles.input} value={name} onChange={(event) => setName(event.target.value)} placeholder="Example: Fort Hope Service" /></Field>
            <Field label="Customer / GC"><input style={styles.input} value={customer} onChange={(event) => setCustomer(event.target.value)} placeholder="Example: Penn-co Construction" /></Field>
            <Field label="Location"><input style={styles.input} value={location} onChange={(event) => setLocation(event.target.value)} placeholder="Example: Fort Hope" /></Field>
            <Field label="Projected Start"><input style={styles.input} type="date" value={startDate} onChange={(event) => setStartDate(event.target.value)} /></Field>
            <Field label="Projected Finish"><input style={styles.input} type="date" value={finishDate} onChange={(event) => setFinishDate(event.target.value)} /></Field>
            <Field label="Labour Budget Allowed ($)"><input style={styles.input} type="number" value={labourBudget} onChange={(event) => setLabourBudget(Number(event.target.value))} placeholder="Example: 92000" /></Field>
            <Field label="Project Budget"><input style={styles.input} type="number" value={budget} onChange={(event) => setBudget(Number(event.target.value))} placeholder="Example: 185000" /></Field>
          </div>
        </Card>

        <Card>
          <div style={styles.compactCardHeader}>
            <div>
              <h3 style={styles.sideTitle}>Assign Employees</h3>
              <p style={styles.muted}>Click employees below to assign them to this job. Only assigned crew/foremen will see the job.</p>
            </div>
            <span style={styles.smallPill}>{crewIds.length} assigned</span>
          </div>
          <div style={styles.employeePickGrid}>
            {assignableUsers.map((user) => {
              const selected = crewIds.includes(user.id);
              return (
                <button key={user.id} type="button" style={selected ? styles.employeePickSelected : styles.employeePick} onClick={() => toggleCrew(user.id)}>
                  <div style={styles.smallAvatar}>{initials(user.name)}</div>
                  <div>
                    <strong>{user.name}</strong>
                    <span>{user.role}</span>
                  </div>
                </button>
              );
            })}
          </div>
        </Card>

        <Card>
          <div style={styles.compactCardHeader}>
            <div>
              <h3 style={styles.sideTitle}>Preloaded PM Checklist</h3>
              <p style={styles.muted}>This checklist will be automatically added to the job file.</p>
            </div>
            <span style={styles.smallPill}>{pmTaskList(0).length} tasks</span>
          </div>
          <div style={styles.checklistPreview}>
            {pmTaskList(0).slice(0, 8).map((task) => <span key={task.id}>{task.title}</span>)}
            <span>+ more checklist items...</span>
          </div>
        </Card>

        <div style={styles.modalActions}>
          <button style={styles.secondary} onClick={onClose}>Cancel</button>
          <button style={styles.primary} onClick={() => onCreate({ name, customer, location, startDate, finishDate, crewIds, labourBudget, budget })}><Plus size={16} /> Create Job</button>
        </div>
      </div>
    </div>
  );
}

function SettingsModal({ users, setUsers, currentUser, onClose }: { users: User[]; setUsers: (users: User[]) => void; currentUser: User; onClose: () => void }) {
  const isAdmin = currentUser.role === "Admin" || currentUser.role === "Project Manager";
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [role, setRole] = useState<Role>("Crew");
  const [monthlyPointLimit, setMonthlyPointLimit] = useState(25);
  const [temporaryPassword, setTemporaryPassword] = useState("temp123");

  function addUser() {
    if (!isAdmin || !name.trim()) return;
    setUsers([
      ...users,
      {
        id: Date.now(),
        name: name.trim(),
        role,
        points: 0,
        monthlyPointLimit: role === "Foreman" ? monthlyPointLimit : 0,
        email: email.trim(),
        password: temporaryPassword || "temp123",
        mustSetPassword: true,
        active: true,
      },
    ]);
    setName("");
    setEmail("");
    setRole("Crew");
    setMonthlyPointLimit(25);
    setTemporaryPassword("temp123");
  }

  function updateUser(id: number, update: Partial<User>) {
    if (!isAdmin) return;
    setUsers(users.map((user) => user.id === id ? { ...user, ...update } : user));
  }

  function deleteUser(id: number) {
    if (!isAdmin) return;
    if (id === currentUser.id) {
      window.alert("You cannot delete the account you are currently signed into.");
      return;
    }
    const targetUser = users.find((user) => user.id === id);
    if (!targetUser) return;
    if (!window.confirm(`Delete user ${targetUser.name}? This cannot be undone in this prototype.`)) return;
    setUsers(users.filter((user) => user.id !== id));
  }

  function resetTemporaryPassword(id: number) {
    const newTempPassword = window.prompt("Enter a new temporary password for this user:", "temp123");
    if (!newTempPassword) return;
    updateUser(id, { password: newTempPassword, mustSetPassword: true });
  }

  return (
    <div style={styles.modalBackdrop}>
      <div style={styles.settingsModal}>
        <div style={styles.cardHeader}>
          <div>
            <h2 style={styles.cardTitle}>Settings</h2>
            <p style={styles.muted}>Admin user management: create accounts, assign roles, set temporary passwords, force password resets, disable accounts, and set foreman point allowance.</p>
          </div>
          <button style={styles.closeButton} onClick={onClose}>×</button>
        </div>

        {!isAdmin ? (
          <Card>
            <h3 style={styles.sideTitle}>Access Restricted</h3>
            <p style={styles.muted}>Only Admin or Project Manager users can change user settings.</p>
          </Card>
        ) : (
          <>
            <Card>
              <h3 style={styles.sideTitle}>Create Employee Account</h3>
              <p style={styles.cardHelp}>Admin creates the account with a temporary password. The employee must set their own password on first login.</p>
              <div style={styles.userFormGridEnhanced}>
                <Field label="Name"><input style={styles.input} value={name} onChange={(event) => setName(event.target.value)} placeholder="Enter crew or foreman name" /></Field>
                <Field label="Email"><input style={styles.input} value={email} onChange={(event) => setEmail(event.target.value)} placeholder="name@pjselectric.ca" /></Field>
                <Field label="Role"><select style={styles.input} value={role} onChange={(event) => setRole(event.target.value as Role)}><option>Crew</option><option>Foreman</option><option>Project Manager</option><option>Admin</option></select></Field>
                <Field label="Temp Password"><input style={styles.input} value={temporaryPassword} onChange={(event) => setTemporaryPassword(event.target.value)} placeholder="Temporary password" /></Field>
                <Field label="Foreman Monthly Point Allowance"><input style={styles.input} type="number" value={monthlyPointLimit} disabled={role !== "Foreman"} onChange={(event) => setMonthlyPointLimit(Number(event.target.value))} /></Field>
                <button style={styles.primary} onClick={addUser}><Plus size={16} /> Create User</button>
              </div>
            </Card>

            <Card>
              <h3 style={styles.sideTitle}>User Accounts</h3>
              <p style={styles.cardHelp}>Use Reset Temp Password if someone forgets their password. This forces them to set a new password next login.</p>
              <div style={styles.userManageList}>
                {users.map((user) => (
                  <div style={styles.userManageRowEnhanced} key={user.id}>
                    <div style={styles.userManageName}>
                      <div style={styles.smallAvatar}>{initials(user.name)}</div>
                      <div>
                        <strong>{user.name}</strong>
                        <span>{user.email || "No email set"}</span>
                        <small>{user.points} pts · {user.mustSetPassword ? "Password reset required" : "Password set"}</small>
                      </div>
                    </div>
                    <select style={styles.input} value={user.role} onChange={(event) => updateUser(user.id, { role: event.target.value as Role, monthlyPointLimit: event.target.value === "Foreman" ? user.monthlyPointLimit || 25 : user.monthlyPointLimit })}><option>Crew</option><option>Foreman</option><option>Project Manager</option><option>Admin</option></select>
                    <input style={styles.input} type="number" value={user.monthlyPointLimit} disabled={user.role !== "Foreman"} onChange={(event) => updateUser(user.id, { monthlyPointLimit: Number(event.target.value) })} />
                    <label style={styles.scheduledCheck}><input type="checkbox" checked={user.active !== false} onChange={(event) => updateUser(user.id, { active: event.target.checked })} /> Active</label>
                    <label style={styles.scheduledCheck}><input type="checkbox" checked={!!user.mustSetPassword} onChange={(event) => updateUser(user.id, { mustSetPassword: event.target.checked })} /> Force Reset</label>
                    <button style={styles.secondary} onClick={() => resetTemporaryPassword(user.id)}>Reset Temp Password</button>
                    <button style={styles.danger} onClick={() => deleteUser(user.id)}><Trash2 size={15} /> Delete</button>
                  </div>
                ))}
              </div>
            </Card>
          </>
        )}
      </div>
    </div>
  );
}

function PMRequestCard({ job, updateJob, toggleTask, isAdmin, currentUser }: { job: Job; updateJob: (update: Partial<Job>) => void; toggleTask: (id: number) => void; isAdmin: boolean; currentUser: User }) {
  const [requestText, setRequestText] = useState("");
  function addRequest() {
    if (!requestText.trim()) return;
    updateJob({ pmRequests: [{ id: Date.now(), title: `${requestText.trim()} — requested by ${currentUser.name}`, done: false }, ...job.pmRequests] });
    setRequestText("");
  }
  return <Card><div style={styles.compactCardHeader}><div><h3 style={styles.sideTitle}>{isAdmin ? "PM Requests From Site" : "Create Task For PM"}</h3><p style={styles.cardHelp}>{isAdmin ? "Requests created by crew/foremen. PM/Admin checks these off after they are handled." : "Add anything you need the PM/Admin to do for this job."}</p></div><span style={styles.smallPill}>{job.pmRequests.filter((task) => !task.done).length} open</span></div>{!isAdmin && <div style={styles.pmRequestForm}><input style={styles.input} value={requestText} onChange={(event) => setRequestText(event.target.value)} placeholder="Example: order material, answer RFI, book inspection, confirm change..." /><button style={styles.smallButton} onClick={addRequest}><Plus size={16} /> Send To PM</button></div>}<div style={styles.cleanTaskList}>{job.pmRequests.map((task) => <label style={task.done ? styles.cleanTaskDone : styles.cleanTask} key={task.id}><input type="checkbox" checked={task.done} disabled={!isAdmin} onChange={() => isAdmin && toggleTask(task.id)} style={styles.cleanCheckbox} /><span>{task.title}</span></label>)}</div></Card>;
}

function TaskCard({ title, tasks, onToggle, canEdit = true, lockedNote }: { title: string; tasks: Task[]; onToggle: (id: number) => void; canEdit?: boolean; lockedNote?: string }) {
  const complete = tasks.filter((task) => task.done).length;
  return <Card><div style={styles.compactCardHeader}><h3 style={styles.sideTitle}>{title}</h3><span style={styles.smallPill}>{complete}/{tasks.length}</span></div>{lockedNote && <p style={styles.lockedNote}>{lockedNote}</p>}<div style={styles.cleanTaskList}>{tasks.map((task) => <label style={task.done ? styles.cleanTaskDone : styles.cleanTask} key={task.id}><input type="checkbox" checked={task.done} disabled={!canEdit} onChange={() => canEdit && onToggle(task.id)} style={styles.cleanCheckbox} /><span>{task.title}</span></label>)}</div></Card>;
}

function InspectionCard({ inspections, canEdit, onChange }: { inspections: Inspection[]; canEdit: boolean; onChange: (id: number, status: string) => void }) {
  return <Card><div style={styles.compactCardHeader}><h3 style={styles.sideTitle}>Inspections</h3><span style={styles.smallPill}>{inspections.length}</span></div><div style={styles.cleanInspectionList}>{inspections.map((inspection) => <div style={styles.cleanInspection} key={inspection.id}><div style={styles.inspectionText}><strong>{inspection.title}</strong><span>{inspection.status}{inspection.date ? ` · ${inspection.date}` : ""}</span><small>{inspection.notes}</small></div><select style={canEdit ? styles.cleanSelect : styles.readOnlySelect} value={inspection.status} disabled={!canEdit} onChange={(event) => onChange(inspection.id, event.target.value)}><option>Not Started</option><option>Pending</option><option>Scheduled</option><option>Passed</option><option>Failed</option></select></div>)}</div></Card>;
}

function DocumentPreviewModal({ document, onClose }: { document: JobDocument; onClose: () => void }) {
  return (
    <div style={styles.modalBackdrop}>
      <div style={styles.previewModal}>
        <div style={styles.cardHeader}>
          <div>
            <h2 style={styles.cardTitle}>{document.fileName}</h2>
            <p style={styles.muted}>{document.type} · uploaded by {document.uploadedBy} · {document.uploadedDate}</p>
          </div>
          <button style={styles.closeButton} onClick={onClose}>×</button>
        </div>
        {document.type === "Photo" && document.url ? (
          <img src={document.url} alt={document.fileName} style={styles.previewImage} />
        ) : document.url ? (
          <iframe src={document.url} title={document.fileName} style={styles.previewFrame} />
        ) : (
          <div style={styles.previewEmpty}>Preview not available for older sample files. New uploaded photos/documents will be viewable here.</div>
        )}
      </div>
    </div>
  );
}

function ModuleCard({ title, icon, items }: { title: string; icon: ReactNode; items: string[] }) {
  return <Card><div style={styles.moduleHeader}>{icon}<h3 style={styles.sideTitle}>{title}</h3></div><div style={styles.moduleList}>{items.map((item) => <div key={item} style={styles.moduleItem}>{item}</div>)}</div></Card>;
}

function PageTitle({ title, subtitle }: { title: string; subtitle: string }) { return <Card><h2 style={styles.cardTitle}>{title}</h2><p style={styles.muted}>{subtitle}</p></Card>; }
function SideButton({ icon, label, active, onClick }: { icon: ReactNode; label: string; active?: boolean; onClick: () => void }) { return <button style={active ? styles.sideActive : styles.sideButton} onClick={onClick}>{icon}<span>{label}</span></button>; }
function Card({ children, style }: { children: ReactNode; style?: CSSProperties }) { return <div style={{ ...styles.card, ...style }}>{children}</div>; }
function Field({ label, children }: { label: string; children: ReactNode }) { return <label style={styles.field}><span>{label}</span>{children}</label>; }
function Detail({ label, value }: { label: string; value: ReactNode }) { return <div style={styles.detail}><span>{label}</span><strong>{value}</strong></div>; }
function StatusItem({ label, value }: { label: string; value: string }) { return <div style={styles.statusTile}><span>{label}</span><strong>{value}</strong></div>; }
function SearchBox({ value, onChange }: { value: string; onChange: (value: string) => void }) { return <div style={styles.searchBox}><Search size={18} /><input style={styles.searchInput} value={value} onChange={(event) => onChange(event.target.value)} placeholder="Search jobs..." /></div>; }
function Badge({ children, tone }: { children: ReactNode; tone: "blue" | "green" | "yellow" | "red" }) { const toneStyle = tone === "blue" ? styles.badgeBlue : tone === "green" ? styles.badgeGreen : tone === "red" ? styles.badgeRed : styles.badgeYellow; return <span style={{ ...styles.badge, ...toneStyle }}>{children}</span>; }
function Bar({ label, value, helper, warn, blue }: { label: string; value: number; helper: string; warn?: boolean; blue?: boolean }) { return <div style={styles.barRow}><div style={styles.barTop}><strong>{label}</strong><span>{helper}</span></div><div style={styles.barOuter}><div style={{ ...styles.barInner, width: `${Math.max(0, Math.min(100, value))}%`, background: warn ? "#f59e0b" : blue ? "#0ea5e9" : "#047857" }} /></div></div>; }
function HealthTile({ title, value, subtitle, percent, warn, blue }: { title: string; value: string; subtitle: string; percent: number; warn?: boolean; blue?: boolean }) { return <Card style={styles.healthTile}><div style={styles.tileTop}><span>{title}</span><strong>{value}</strong></div><div style={styles.barOuter}><div style={{ ...styles.barInner, width: `${percent}%`, background: warn ? "#f59e0b" : blue ? "#0ea5e9" : "#047857" }} /></div><small style={styles.tileSub}>{subtitle}</small></Card>; }
function MiniHealth({ job }: { job: Job }) { const labour = pct(job.labourCostToDate || 0, job.labourBudget || 0); return <div style={styles.miniHealth}><span>Progress {job.progress}%</span><div style={styles.miniBar}><i style={{ ...styles.miniFill, width: `${job.progress}%` }} /></div><span>Labour {labour}%</span><div style={styles.miniBar}><i style={{ ...styles.miniFill, width: `${labour}%`, background: labour > 85 ? "#f59e0b" : "#047857" }} /></div></div>; }
function CrewCard({ crew }: { crew: User[] }) { return <Card style={styles.crewCard}><h3 style={styles.sideTitle}>Crew Assignments ({crew.length})</h3><strong style={styles.statValue}>{crew.length}</strong><div style={styles.chips}>{crew.map((person) => <span style={styles.chip} key={person.id}>{person.name}</span>)}</div></Card>; }

function MobileCrewStyles() {
  return (
    <style>{`
      @media (max-width: 820px) {
        #jobflow-app {
          display: block !important;
          min-height: 100vh !important;
          padding-bottom: 86px !important;
          background: #eef3f8 !important;
        }

        #jobflow-app > aside {
          position: fixed !important;
          left: 0 !important;
          right: 0 !important;
          bottom: 0 !important;
          z-index: 999 !important;
          height: 76px !important;
          padding: 8px 10px !important;
          display: flex !important;
          flex-direction: row !important;
          align-items: center !important;
          gap: 8px !important;
          overflow-x: auto !important;
          overflow-y: hidden !important;
          border-top: 1px solid rgba(255,255,255,.18) !important;
          box-shadow: 0 -10px 24px rgba(15,23,42,.24) !important;
        }

        #jobflow-app > aside > div:first-child,
        #jobflow-app > aside > h2,
        #jobflow-app > aside > div:last-child {
          display: none !important;
        }

        #jobflow-app > aside button {
          min-width: 92px !important;
          height: 58px !important;
          padding: 8px !important;
          border-radius: 14px !important;
          display: grid !important;
          place-items: center !important;
          gap: 4px !important;
          font-size: 11px !important;
          line-height: 1.05 !important;
          flex: 0 0 auto !important;
        }

        #jobflow-app main {
          padding: 10px !important;
          width: 100% !important;
          min-width: 0 !important;
          overflow-x: hidden !important;
        }

        #jobflow-app header {
          display: grid !important;
          grid-template-columns: 1fr !important;
          gap: 10px !important;
          padding: 14px !important;
          border-radius: 14px !important;
          margin-bottom: 10px !important;
        }

        #jobflow-app header h1 {
          font-size: 20px !important;
          margin: 4px 0 !important;
        }

        #jobflow-app header p {
          display: none !important;
        }

        #jobflow-app header > div:last-child {
          display: grid !important;
          grid-template-columns: 1fr !important;
          gap: 8px !important;
        }

        #jobflow-app section,
        #jobflow-app [style*="grid-template-columns"] {
          grid-template-columns: 1fr !important;
        }

        #jobflow-app button,
        #jobflow-app input,
        #jobflow-app select,
        #jobflow-app textarea,
        #jobflow-app label {
          font-size: 16px !important;
        }

        #jobflow-app textarea {
          min-height: 140px !important;
        }

        #jobflow-app [style*="position: fixed"] {
          max-width: 100vw !important;
        }

        #jobflow-app [style*="minmax(390px"],
        #jobflow-app [style*="minmax(380px"],
        #jobflow-app [style*="minmax(420px"] {
          grid-template-columns: 1fr !important;
        }
      }

      @media (min-width: 821px) and (max-width: 1180px) {
        #jobflow-app {
          grid-template-columns: 82px minmax(0, 1fr) !important;
        }

        #jobflow-app > aside {
          padding: 12px 8px !important;
          align-items: stretch !important;
        }

        #jobflow-app > aside > div:first-child,
        #jobflow-app > aside > h2,
        #jobflow-app > aside > div:last-child {
          display: none !important;
        }

        #jobflow-app > aside button {
          min-height: 62px !important;
          padding: 8px 6px !important;
          display: grid !important;
          place-items: center !important;
          gap: 4px !important;
          font-size: 10px !important;
          text-align: center !important;
          line-height: 1.05 !important;
        }

        #jobflow-app main {
          padding: 12px !important;
          overflow-x: hidden !important;
        }

        #jobflow-app header {
          display: grid !important;
          grid-template-columns: 1fr !important;
          gap: 12px !important;
          padding: 18px !important;
        }

        #jobflow-app header > div:last-child {
          display: grid !important;
          grid-template-columns: minmax(180px, 1fr) minmax(260px, 1fr) auto !important;
          gap: 10px !important;
        }

        #jobflow-app section,
        #jobflow-app [style*="grid-template-columns: minmax(0, 0.9fr)"],
        #jobflow-app [style*="grid-template-columns: minmax(0, 1.45fr)"],
        #jobflow-app [style*="grid-template-columns: repeat(4"],
        #jobflow-app [style*="grid-template-columns: repeat(7"],
        #jobflow-app [style*="grid-template-columns: 1.35fr"] {
          grid-template-columns: 1fr !important;
        }

        #jobflow-app [style*="grid-template-columns: 1.35fr"] {
          overflow-x: auto !important;
        }

        #jobflow-app button,
        #jobflow-app input,
        #jobflow-app select,
        #jobflow-app textarea {
          max-width: 100% !important;
        }
      }
    `}</style>
  );
}

const styles: Record<string, CSSProperties> = {
  loginShell: { minHeight: "100vh", display: "grid", placeItems: "center", background: "linear-gradient(135deg,#07182d,#0f172a)", padding: 18, fontFamily: "Inter, Arial, sans-serif" },
  loginCard: { width: "min(460px, 100%)", background: "white", borderRadius: 22, padding: 26, display: "grid", gap: 18, boxShadow: "0 30px 90px rgba(0,0,0,.28)", color: "#0f172a" },
  loginLogo: { background: "#00a86b", color: "white", borderRadius: 16, padding: 20, textAlign: "center", fontWeight: 900, fontSize: 25, lineHeight: 1.05 },
  loginTitle: { margin: "0 0 6px", fontSize: 30, color: "#0f172a", fontWeight: 900 },
  loginButton: { border: 0, background: "#047857", color: "white", borderRadius: 12, padding: "14px 16px", fontWeight: 900, cursor: "pointer", fontSize: 16 },
  loginNote: { margin: 0, color: "#64748b", fontWeight: 700, fontSize: 12, lineHeight: 1.45 },
  loginError: { background: "#fee2e2", color: "#991b1b", border: "1px solid #fecaca", borderRadius: 10, padding: 10, fontWeight: 800, fontSize: 13 },
  app: { minHeight: "100vh", display: "grid", gridTemplateColumns: "230px minmax(0, 1fr)", background: "#eef3f8", color: "#0f172a", fontFamily: "Inter, Arial, sans-serif" },
  sidebar: { background: "linear-gradient(180deg,#07182d,#0f172a)", color: "white", padding: 16, display: "flex", flexDirection: "column", gap: 14 },
  logo: { background: "#00a86b", borderRadius: 12, padding: 18, textAlign: "center", fontWeight: 900, fontSize: 22, lineHeight: 1.1 },
  sidebarTitle: { fontSize: 18, margin: "10px 6px 18px", fontWeight: 900 },
  sideButton: { border: 0, background: "transparent", color: "#e2e8f0", padding: "13px 12px", borderRadius: 10, display: "flex", alignItems: "center", gap: 12, fontSize: 15, fontWeight: 700, cursor: "pointer" },
  sideActive: { border: 0, background: "#00a86b", color: "white", padding: "13px 12px", borderRadius: 10, display: "flex", alignItems: "center", gap: 12, fontSize: 15, fontWeight: 800, cursor: "pointer" },
  userCard: { marginTop: "auto", border: "1px solid #334155", borderRadius: 14, padding: 12, display: "grid", gap: 10 },
  userTop: { display: "flex", alignItems: "center", gap: 10 },
  userText: { display: "grid", gap: 2 },
  avatar: { width: 44, height: 44, borderRadius: 10, background: "#00a86b", display: "grid", placeItems: "center", fontWeight: 900 },
  signOut: { border: 0, background: "transparent", color: "white", display: "flex", alignItems: "center", gap: 8, fontWeight: 700, cursor: "pointer" },
  main: { padding: 16, overflowY: "auto", overflowX: "hidden", minWidth: 0, width: "100%", maxWidth: "100%", boxSizing: "border-box" },
  header: { background: "white", border: "1px solid #dbe3ed", borderRadius: 18, padding: "22px 28px", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 18, marginBottom: 16, boxShadow: "0 8px 20px rgba(15,23,42,.05)", minWidth: 0 },
  brand: { color: "#047857", letterSpacing: 3, fontWeight: 900, fontSize: 13 },
  title: { margin: "10px 0 6px", fontSize: 26, color: "#0f172a", fontWeight: 900 },
  subtitle: { margin: 0, color: "#334155", fontWeight: 500 },
  headerActions: { display: "grid", gridTemplateColumns: "auto minmax(220px, 1fr) auto", gap: 10, alignItems: "center", minWidth: 0 },
  headerActionsLocked: { display: "grid", gridTemplateColumns: "auto minmax(180px, 1fr) auto", gap: 10, alignItems: "center", minWidth: 0 },
  currentUserBadge: { border: "1px solid #cbd5e1", borderRadius: 10, background: "#f8fafc", color: "#0f172a", padding: "8px 12px", display: "grid", gap: 2, fontWeight: 800 },
  signedIn: { fontSize: 13, fontWeight: 900 },
  selectWide: { border: "1px solid #cbd5e1", borderRadius: 8, padding: "9px 10px", color: "#0f172a", fontWeight: 700, background: "white" },
  primary: { border: 0, background: "#047857", color: "white", borderRadius: 9, padding: "10px 14px", display: "inline-flex", alignItems: "center", justifyContent: "center", gap: 8, fontWeight: 900, cursor: "pointer" },
  secondary: { border: "1px solid #cbd5e1", background: "white", color: "#0f172a", borderRadius: 9, padding: "10px 14px", display: "inline-flex", alignItems: "center", justifyContent: "center", gap: 8, fontWeight: 800, cursor: "pointer" },
  dashboardGrid: { display: "grid", gridTemplateColumns: "minmax(0, 1.45fr) minmax(330px, 0.8fr)", gap: 14, alignItems: "start", width: "100%" },
  jobFileHeader: { display: "grid", gridTemplateColumns: "auto 1fr auto", gap: 16, alignItems: "center", background: "white", border: "1px solid #d5dde8", borderRadius: 15, padding: 16, boxShadow: "0 8px 20px rgba(15,23,42,.04)" },
  jobFileTitle: { margin: 0, color: "#0f172a", fontSize: 28, fontWeight: 900 },
  jobFileMainGrid: { display: "grid", gridTemplateColumns: "minmax(0, 0.9fr) minmax(0, 1.1fr)", gap: 14, alignItems: "start", width: "100%", minWidth: 0 },
  bottomGrid: { display: "grid", gridTemplateColumns: "repeat(4, minmax(210px, 1fr))", gap: 12, marginTop: 14, alignItems: "start", width: "100%" },
  healthTiles: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(145px, 1fr))", gap: 10, marginBottom: 14, width: "100%", minWidth: 0 },
  healthTile: { minHeight: 135, display: "grid", gridTemplateRows: "auto auto 1fr", gap: 10, alignItems: "center", minWidth: 0 },
  tileTop: { display: "grid", gridTemplateColumns: "1fr", justifyItems: "center", alignItems: "center", gap: 6, color: "#0f172a", fontWeight: 900, textAlign: "center", lineHeight: 1.15 },
  tileSub: { color: "#334155", fontWeight: 700, textAlign: "center", lineHeight: 1.35, display: "block" },
  card: { background: "white", border: "1px solid #d5dde8", borderRadius: 15, padding: 14, boxShadow: "0 8px 20px rgba(15,23,42,.04)", color: "#0f172a", minWidth: 0, overflow: "hidden", boxSizing: "border-box" },
  cardHeader: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 16, marginBottom: 14, minWidth: 0 },
  compactCardHeader: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 10, marginBottom: 12, minWidth: 0 },
  cardTitle: { margin: 0, fontSize: 22, fontWeight: 900, color: "#0f172a" },
  muted: { color: "#334155", fontWeight: 600, fontSize: 14 },
  sideTitle: { margin: 0, textAlign: "left", color: "#0f172a", fontSize: 18, fontWeight: 900 },
  searchBox: { border: "1px solid #cbd5e1", borderRadius: 9, padding: "9px 12px", display: "flex", alignItems: "center", gap: 10, background: "white" },
  searchInput: { border: 0, outline: 0, color: "#0f172a", fontWeight: 700, minWidth: 150 },
  jobList: { border: "1px solid #cbd5e1", borderRadius: 12, overflow: "hidden" },
  jobListHeader: { display: "grid", gridTemplateColumns: "1.35fr 80px 120px 70px 145px 130px", gap: 8, padding: "12px 14px", background: "#f1f5f9", color: "#0f172a", fontSize: 12, fontWeight: 900, textTransform: "uppercase" },
  jobRow: { width: "100%", border: 0, borderTop: "1px solid #cbd5e1", display: "grid", gridTemplateColumns: "1.35fr 80px 120px 70px 145px 130px", gap: 8, alignItems: "center", textAlign: "left", padding: "14px", cursor: "pointer", minWidth: 0 },
  jobName: { color: "#047857", fontWeight: 900, display: "block", marginBottom: 6 },
  jobSub: { color: "#0f172a", fontWeight: 600, fontSize: 12, display: "block", lineHeight: 1.4 },
  tableText: { color: "#0f172a", fontWeight: 700, fontSize: 13 },
  badge: { borderRadius: 999, padding: "7px 10px", fontSize: 12, fontWeight: 900, textAlign: "center", display: "inline-block" },
  badgeBlue: { background: "#dbeafe", color: "#1d4ed8" },
  badgeGreen: { background: "#dcfce7", color: "#166534" },
  badgeYellow: { background: "#fef3c7", color: "#92400e" },
  badgeRed: { background: "#fee2e2", color: "#991b1b" },
  miniHealth: { display: "grid", gap: 3, color: "#334155", fontSize: 11, fontWeight: 800 },
  miniBar: { height: 6, background: "#e2e8f0", borderRadius: 999, overflow: "hidden" },
  miniFill: { display: "block", height: "100%", background: "#047857", borderRadius: 999 },
  crewCard: { minHeight: 160, display: "grid", placeItems: "center", gap: 8, textAlign: "center" },
  statValue: { fontSize: 28, color: "#0f172a", fontWeight: 900, lineHeight: 1.1, textAlign: "center" },
  chips: { display: "flex", flexWrap: "wrap", gap: 8, justifyContent: "center" },
  chip: { background: "#dcfce7", color: "#166534", border: "1px solid #86efac", borderRadius: 999, padding: "7px 12px", fontWeight: 800, fontSize: 13 },
  jobDetailTitle: { textAlign: "center", fontSize: 24, color: "#0f172a", fontWeight: 900, margin: "8px 0 4px", lineHeight: 1.15 },
  jobDetailSub: { textAlign: "center", color: "#1e293b", fontWeight: 700, marginBottom: 14 },
  detailGrid: { display: "grid", gridTemplateColumns: "1fr 1fr", border: "1px solid #cbd5e1", borderRadius: 12, overflow: "hidden", marginBottom: 14 },
  detail: { padding: 12, borderRight: "1px solid #e2e8f0", borderBottom: "1px solid #e2e8f0", display: "grid", gap: 6, textAlign: "center" },
  statusPanel: { border: "1px solid #cbd5e1", borderRadius: 12, padding: 14, background: "#f8fafc", marginBottom: 14 },
  statusEditBox: { border: "1px solid #bbf7d0", borderRadius: 12, padding: 14, background: "#f0fdf4", marginBottom: 14 },
  statusEditGrid: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 12 },
  folderTabs: { display: "grid", gridTemplateColumns: "repeat(4, minmax(0, 1fr))", gap: 8, margin: "14px 0" },
  jobFolderTabs: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(95px, 1fr))", gap: 8, margin: "14px 0", width: "100%", minWidth: 0 },
  folderTab: { border: "1px solid #cbd5e1", background: "white", color: "#0f172a", borderRadius: 9, padding: "9px 6px", fontWeight: 900, cursor: "pointer", fontSize: 11, minHeight: 56, display: "flex", alignItems: "center", justifyContent: "center", textAlign: "center", lineHeight: 1.15, whiteSpace: "normal", minWidth: 0 },
  folderTabActive: { border: "1px solid #047857", background: "#047857", color: "white", borderRadius: 9, padding: "9px 6px", fontWeight: 900, cursor: "pointer", fontSize: 11, minHeight: 56, display: "flex", alignItems: "center", justifyContent: "center", textAlign: "center", lineHeight: 1.15, whiteSpace: "normal", minWidth: 0 },
  folderPanel: { border: "1px solid #dbe3ed", borderRadius: 12, padding: 14, background: "#ffffff", marginTop: 10, display: "grid", gap: 12 },
  statusGrid: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 },
  statusTile: { background: "white", border: "1px solid #dbe3ed", borderRadius: 10, padding: 12, display: "grid", gap: 5, color: "#0f172a" },
  healthPanel: { border: "1px solid #cbd5e1", borderRadius: 12, padding: 14, background: "#f8fafc", marginBottom: 14 },
  barRow: { display: "grid", gap: 6, marginTop: 10 },
  barTop: { display: "grid", gridTemplateColumns: "minmax(0,1fr) auto", gap: 10, color: "#0f172a", fontSize: 13, fontWeight: 700, alignItems: "center" },
  barOuter: { width: "100%", height: 10, background: "#e2e8f0", borderRadius: 999, overflow: "hidden" },
  barInner: { height: "100%", borderRadius: 999 },
  nextActionBox: { border: "1px solid #cbd5e1", borderRadius: 12, padding: 14, background: "#f8fafc", marginBottom: 14 },
  nextActionTop: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, color: "#334155", fontWeight: 800, fontSize: 12 },
  nextActionInput: { width: "100%", minHeight: 88, border: "1px solid #10b981", borderRadius: 10, padding: 12, boxSizing: "border-box", background: "white", color: "#0f172a", fontWeight: 700, lineHeight: 1.45, resize: "vertical", marginTop: 10 },
  sectionTitle: { margin: 0, color: "#0f172a", fontSize: 18, fontWeight: 900 },
  editBox: { border: "1px solid #d5dde8", borderRadius: 12, padding: 14, background: "white" },
  editGrid: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 12 },
  field: { display: "grid", gap: 6, color: "#0f172a", fontWeight: 900, fontSize: 12 },
  input: { border: "1px solid #cbd5e1", borderRadius: 8, padding: "9px 10px", background: "white", color: "#0f172a", fontWeight: 700, minWidth: 0, boxSizing: "border-box" },
  cleanTaskList: { display: "grid", gap: 8, maxHeight: 520, overflow: "auto", paddingRight: 4 },
  cleanTask: { display: "grid", gridTemplateColumns: "22px 1fr", gap: 10, alignItems: "center", background: "#f8fafc", border: "1px solid #dbe3ed", borderRadius: 12, padding: "12px 14px", color: "#0f172a", fontSize: 15, fontWeight: 800, lineHeight: 1.3 },
  cleanTaskDone: { display: "grid", gridTemplateColumns: "22px 1fr", gap: 10, alignItems: "center", background: "#ecfdf5", border: "1px solid #a7f3d0", borderRadius: 12, padding: "12px 14px", color: "#334155", fontSize: 15, fontWeight: 800, lineHeight: 1.3, textDecoration: "line-through" },
  cleanCheckbox: { width: 16, height: 16, margin: 0 },
  lockedNote: { margin: "-4px 0 10px", color: "#64748b", fontWeight: 800, fontSize: 12 },
  cardHelp: { margin: "4px 0 0", color: "#64748b", fontSize: 12, fontWeight: 700 },
  pmRequestForm: { display: "grid", gridTemplateColumns: "1fr auto", gap: 8, marginBottom: 12 },
  cleanInspectionList: { display: "grid", gap: 10 },
  cleanInspection: { display: "grid", gridTemplateColumns: "1fr", gap: 8, alignItems: "start", background: "#f8fafc", border: "1px solid #dbe3ed", borderRadius: 12, padding: "12px 14px", minWidth: 0 },
  inspectionText: { display: "grid", gridTemplateColumns: "1fr", gap: 4, color: "#0f172a", minWidth: 0, lineHeight: 1.25 },
  cleanSelect: { width: "100%", border: "1px solid #cbd5e1", borderRadius: 8, padding: "8px 9px", background: "white", color: "#0f172a", fontWeight: 800, boxSizing: "border-box" },
  readOnlySelect: { width: "100%", border: "1px solid #dbe3ed", borderRadius: 8, padding: "8px 9px", background: "#f1f5f9", color: "#334155", fontWeight: 900, boxSizing: "border-box", opacity: 1, cursor: "not-allowed" },
  smallPill: { background: "#ecfdf5", color: "#047857", border: "1px solid #a7f3d0", borderRadius: 999, padding: "5px 10px", fontSize: 12, fontWeight: 900, whiteSpace: "nowrap" },
  smallButton: { border: "1px solid #cbd5e1", background: "white", color: "#0f172a", borderRadius: 8, padding: "10px", display: "inline-flex", justifyContent: "center", alignItems: "center", gap: 8, fontWeight: 800, cursor: "pointer" },
  danger: { border: "1px solid #fecaca", background: "#fff1f2", color: "#b91c1c", borderRadius: 9, padding: "9px 12px", display: "inline-flex", alignItems: "center", gap: 6, fontWeight: 900, cursor: "pointer" },
  cleanStack: { display: "grid", gap: 14 },
  moduleGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 14 },
  wideModule: { gridColumn: "span 2" },
  moduleHeader: { display: "flex", alignItems: "center", gap: 10, marginBottom: 12, color: "#047857" },
  moduleList: { display: "grid", gap: 8 },
  moduleItem: { background: "#f8fafc", border: "1px solid #dbe3ed", borderRadius: 10, padding: "10px 12px", fontWeight: 700, color: "#0f172a" },
  potentialItem: { display: "grid", gap: 10, padding: 12, border: "1px solid #dbe3ed", borderRadius: 12, background: "#f8fafc", marginTop: 10 },
  potentialJobFormGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", gap: 10 },
  potentialJobGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 12, marginTop: 12 },
  potentialJobCard: { display: "grid", gap: 8, padding: 10, border: "1px solid #dbe3ed", borderRadius: 10, background: "#f8fafc", color: "#0f172a" },
  potentialJobStats: { display: "grid", gridTemplateColumns: "1fr 1fr", border: "1px solid #cbd5e1", borderRadius: 10, overflow: "hidden", fontSize: 12 },
  potentialActivityBox: { border: "1px solid #dbe3ed", background: "white", borderRadius: 10, padding: 10, display: "grid", gap: 8, fontSize: 12 },
  potentialActivityItem: { borderTop: "1px solid #e2e8f0", paddingTop: 7, display: "grid", gap: 3, color: "#0f172a" },
  potentialSmallGrid: { display: "grid", gridTemplateColumns: "1fr 90px 120px", gap: 10 },
  warning: { background: "#fffbeb", border: "1px solid #fde68a", color: "#92400e", borderRadius: 12, padding: 12, fontWeight: 700 },
  labourGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(380px, 1fr))", gap: 14 },
  pointsForm: { display: "grid", gridTemplateColumns: "1.2fr 120px 2fr auto", gap: 12, alignItems: "end", marginTop: 16 },
  pointsFormWithJob: { display: "grid", gridTemplateColumns: "1.2fr 1.4fr 100px 2fr auto", gap: 12, alignItems: "end", marginTop: 16 },
  limitRows: { display: "grid", gap: 12, marginTop: 14 },
  limitRow: { display: "grid", gridTemplateColumns: "1fr 150px", gap: 14, alignItems: "center", border: "1px solid #d5dde8", borderRadius: 12, padding: 14 },
  pointsGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(250px, 1fr))", gap: 12, marginTop: 14 },
  pointCard: { border: "1px solid #d5dde8", borderRadius: 12, padding: 14, background: "#f8fafc", display: "grid", gap: 10, color: "#0f172a" },
  pointTop: { display: "flex", justifyContent: "space-between", gap: 10 },
  pointHistory: { borderTop: "1px solid #d5dde8", paddingTop: 10, display: "grid", gap: 6, fontWeight: 650 },
  assignedTaskGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(390px, 1fr))", gap: 14 },
  dailyReportBox: { borderTop: "1px solid #dbe3ed", marginTop: 14, paddingTop: 14, display: "grid", gap: 10 },
  dailyReportInput: { width: "100%", minHeight: 100, border: "1px solid #cbd5e1", borderRadius: 10, padding: 12, boxSizing: "border-box", background: "white", color: "#0f172a", fontWeight: 700, resize: "vertical" },
  reportHistoryList: { display: "grid", gap: 8, marginTop: 4 },
  reportHistoryItem: { background: "#f8fafc", border: "1px solid #dbe3ed", borderRadius: 10, padding: 10, display: "grid", gap: 4, color: "#0f172a", fontSize: 13 },
  calendarDateGrid: { display: "flex", flexWrap: "wrap", gap: 8 },
  calendarDate: { border: "1px solid #cbd5e1", background: "white", color: "#0f172a", borderRadius: 9, padding: "8px 10px", fontWeight: 800, cursor: "pointer" },
  calendarDateActive: { border: "1px solid #047857", background: "#047857", color: "white", borderRadius: 9, padding: "8px 10px", fontWeight: 900, cursor: "pointer" },
  calendarEmpty: { color: "#64748b", fontWeight: 800, padding: "8px 0" },
  calendarControls: { display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" },
  calendarLegend: { display: "flex", gap: 10, flexWrap: "wrap", marginTop: 8 },
  awardedLegend: { background: "#ecfdf5", border: "1px solid #047857", color: "#047857", borderRadius: 999, padding: "7px 12px", fontWeight: 900 },
  potentialLegend: { background: "rgba(59,130,246,.12)", border: "1px dashed #3b82f6", color: "#1d4ed8", borderRadius: 999, padding: "7px 12px", fontWeight: 900 },
  calendarGrid: { display: "grid", gridTemplateColumns: "repeat(7, minmax(0, 1fr))", gap: 8 },
  calendarHeaderCell: { background: "#0f172a", color: "white", borderRadius: 10, padding: "10px 8px", textAlign: "center", fontWeight: 900 },
  calendarCell: { minHeight: 145, background: "white", border: "1px solid #dbe3ed", borderRadius: 12, padding: 10, display: "grid", gridTemplateRows: "auto 1fr", gap: 8, minWidth: 0 },
  calendarCellMuted: { minHeight: 145, background: "#f1f5f9", border: "1px solid #e2e8f0", borderRadius: 12, padding: 10, display: "grid", gridTemplateRows: "auto 1fr", gap: 8, opacity: .55, minWidth: 0 },
  calendarDay: { color: "#0f172a", fontSize: 14 },
  calendarItems: { display: "grid", alignContent: "start", gap: 6, minWidth: 0 },
  calendarJobItem: { background: "#ecfdf5", border: "1px solid #bbf7d0", borderLeft: "5px solid #047857", color: "#0f172a", borderRadius: 8, padding: "7px 8px", fontSize: 12, fontWeight: 900, lineHeight: 1.2, overflow: "hidden", textOverflow: "ellipsis" },
  calendarPotentialItem: { background: "rgba(59,130,246,.12)", border: "1px dashed #3b82f6", color: "#1d4ed8", borderRadius: 8, padding: "7px 8px", fontSize: 12, fontWeight: 900, lineHeight: 1.2, overflow: "hidden", textOverflow: "ellipsis" },
  picturePlaceholder: { border: "1px dashed #cbd5e1", background: "white", color: "#64748b", borderRadius: 10, padding: 10, fontWeight: 800, fontSize: 12, display: "flex", flexWrap: "wrap", gap: 8 },
  scopeBox: {
    border: "1px solid #dbe3ed",
    background: "#f8fafc",
    borderRadius: 10,
    padding: 10,
    color: "#0f172a",
    lineHeight: 1.3,
    fontSize: 12,
    maxHeight: 70,
    overflow: "auto",
  },
  picturePreviewButton: { border: "1px solid #cbd5e1", background: "#f8fafc", color: "#0f172a", borderRadius: 999, padding: "7px 10px", fontWeight: 900, cursor: "pointer" },
  alertJobCard: { border: "2px solid #f59e0b", background: "#fffbeb" },
  alertNotice: { background: "#fef3c7", border: "1px solid #f59e0b", color: "#92400e", borderRadius: 10, padding: "8px 10px", fontSize: 12, fontWeight: 900 },
  scheduledCheck: { display: "inline-flex", alignItems: "center", gap: 8, border: "1px solid #cbd5e1", background: "white", color: "#0f172a", borderRadius: 9, padding: "9px 10px", fontWeight: 900, cursor: "pointer", fontSize: 12 },
  smallJobActions: { display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" },
  smallJobsSplitLayout: { display: "grid", gridTemplateColumns: "minmax(320px, .8fr) minmax(420px, 1.2fr)", gap: 12, alignItems: "start" },
  smallJobsList: { display: "grid", gap: 7, maxHeight: "70vh", overflow: "auto", paddingRight: 4 },
  smallJobRow: { border: "1px solid #dbe3ed", background: "#f8fafc", color: "#0f172a", borderRadius: 10, padding: 9, display: "grid", gap: 4, textAlign: "left", cursor: "pointer", fontSize: 12 },
  smallJobRowActive: { border: "1px solid #047857", background: "#ecfdf5", color: "#0f172a", borderRadius: 10, padding: 9, display: "grid", gap: 4, textAlign: "left", cursor: "pointer", fontSize: 12 },
  smallJobRowTitle: { display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center" },
  smallJobRowMeta: { display: "flex", flexWrap: "wrap", gap: 6, alignItems: "center", color: "#475569", fontWeight: 800 },
  smallJobAlertDot: { background: "#f59e0b", color: "white", borderRadius: 999, width: 20, height: 20, display: "inline-grid", placeItems: "center", fontWeight: 900 },
  dailyNoteComposer: { border: "1px solid #bbf7d0", background: "#f0fdf4", borderRadius: 12, padding: 14, display: "grid", gap: 10 },
  pictureList: { display: "flex", flexWrap: "wrap", gap: 8, color: "#0f172a", fontWeight: 800 },
  pictureChip: { background: "white", border: "1px solid #cbd5e1", borderRadius: 999, padding: "6px 10px" },
  addInlineBox: { border: "1px solid #bbf7d0", background: "#f0fdf4", borderRadius: 12, padding: 14, display: "grid", gap: 10 },
  activityList: { display: "grid", gap: 10 },
  activityItem: { background: "#f8fafc", border: "1px solid #dbe3ed", borderRadius: 12, padding: 12, display: "grid", gap: 5, color: "#0f172a" },
  activityType: { width: "fit-content", background: "#dbeafe", color: "#1d4ed8", borderRadius: 999, padding: "5px 9px", fontWeight: 900, fontSize: 12 },
  uploadStatsGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 10 },
  uploadStat: { background: "#f8fafc", border: "1px solid #dbe3ed", borderRadius: 12, padding: 12, display: "grid", gap: 6, color: "#0f172a", minWidth: 0, lineHeight: 1.25 },
  successNotice: { background: "#ecfdf5", border: "1px solid #86efac", color: "#166534", borderRadius: 12, padding: 12, fontWeight: 800 },
  knowifyUploadPanel: { border: "1px solid #bbf7d0", background: "#f0fdf4", borderRadius: 12, padding: 14, display: "grid", gap: 12 },
  uploadFormGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))", gap: 10 },
  dropZone: { border: "2px dashed #94a3b8", background: "white", borderRadius: 14, padding: 26, minHeight: 120, display: "grid", placeItems: "center", gap: 6, textAlign: "center", color: "#0f172a", fontWeight: 800, cursor: "pointer" },
  hiddenFileInput: { display: "none" },
  closeMiniButton: { width: 34, height: 34, border: "1px solid #cbd5e1", borderRadius: 8, background: "white", color: "#0f172a", fontSize: 22, fontWeight: 900, cursor: "pointer" },
  documentFolderList: { display: "grid", gap: 14 },
  documentDateFolder: { border: "1px solid #dbe3ed", background: "#f8fafc", borderRadius: 12, padding: 12, display: "grid", gap: 10 },
  documentGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 10 },
  documentCard: { background: "white", border: "1px solid #dbe3ed", borderRadius: 12, padding: 12, display: "grid", gap: 6, color: "#0f172a", minHeight: 120, textAlign: "left", cursor: "pointer", font: "inherit" },
  documentIcon: { fontSize: 28 },
  documentThumb: { width: "100%", height: 120, objectFit: "cover", borderRadius: 10, border: "1px solid #dbe3ed" },
  previewModal: { background: "white", borderRadius: 18, padding: 24, width: "min(1000px, 95vw)", maxHeight: "92vh", overflow: "auto", boxShadow: "0 30px 90px rgba(0,0,0,.25)", color: "#0f172a" },
  previewImage: { width: "100%", maxHeight: "72vh", objectFit: "contain", borderRadius: 12, background: "#f8fafc", border: "1px solid #dbe3ed" },
  previewFrame: { width: "100%", height: "72vh", border: "1px solid #dbe3ed", borderRadius: 12 },
  previewEmpty: { border: "1px dashed #cbd5e1", borderRadius: 12, padding: 28, textAlign: "center", color: "#334155", fontWeight: 800, background: "#f8fafc" },
  allowanceTop: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 16, marginBottom: 14 },
  allowanceBadge: { background: "#047857", color: "white", borderRadius: 12, padding: "12px 16px", fontWeight: 900, fontSize: 18, whiteSpace: "nowrap" },
  modalBackdrop: { position: "fixed", inset: 0, background: "rgba(15,23,42,.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 24 },
  settingsModal: { background: "white", borderRadius: 18, padding: 24, width: "min(1100px, 95vw)", maxHeight: "88vh", overflow: "auto", boxShadow: "0 30px 90px rgba(0,0,0,.25)", color: "#0f172a" },
  closeButton: { width: 42, height: 42, border: "1px solid #cbd5e1", borderRadius: 10, background: "#f8fafc", color: "#0f172a", fontSize: 26, fontWeight: 900, cursor: "pointer" },
  userFormGrid: { display: "grid", gridTemplateColumns: "1.4fr 180px 220px auto", gap: 12, alignItems: "end", marginTop: 12 },
  userFormGridEnhanced: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 12, alignItems: "end", marginTop: 12 },
  addJobGrid: { display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 12 },
  employeePickGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 10 },
  employeePick: { border: "1px solid #dbe3ed", background: "#f8fafc", borderRadius: 12, padding: 12, display: "flex", alignItems: "center", gap: 10, textAlign: "left", color: "#0f172a", cursor: "pointer" },
  employeePickSelected: { border: "1px solid #047857", background: "#ecfdf5", borderRadius: 12, padding: 12, display: "flex", alignItems: "center", gap: 10, textAlign: "left", color: "#0f172a", cursor: "pointer" },
  checklistPreview: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 8 },
  modalActions: { display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 14 },
  userManageList: { display: "grid", gap: 10, marginTop: 12 },
  userManageRow: { display: "grid", gridTemplateColumns: "1.4fr 180px 180px", gap: 12, alignItems: "center", border: "1px solid #dbe3ed", background: "#f8fafc", borderRadius: 12, padding: 12 },
  userManageRowEnhanced: { display: "grid", gridTemplateColumns: "minmax(220px, 1.5fr) 150px 130px 105px 120px auto auto", gap: 10, alignItems: "center", border: "1px solid #dbe3ed", background: "#f8fafc", borderRadius: 12, padding: 12 },
  userManageName: { display: "flex", alignItems: "center", gap: 10, minWidth: 0 },
  smallAvatar: { width: 38, height: 38, borderRadius: 10, background: "#047857", color: "white", display: "grid", placeItems: "center", fontWeight: 900 },
};
