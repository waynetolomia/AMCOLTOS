import React, { useEffect, useState } from "react";
import { Download, FileText, Plus, Printer, Save, Trash2, ClipboardList, GraduationCap, RefreshCw, Sparkles, Key, Settings } from "lucide-react";
import { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, BorderStyle, TextRun, AlignmentType, HeadingLevel } from "docx";

const SAMPLE_OBJECTIVES = [
  { id: crypto.randomUUID(), text: "Explain basic concepts, terms, and principles related to the course.", hours: 6 },
  { id: crypto.randomUUID(), text: "Apply course knowledge to solve realistic maritime or vocational situations.", hours: 9 },
  { id: crypto.randomUUID(), text: "Analyze problems and select appropriate procedures or solutions.", hours: 9 },
];

function classNames(...items) {
  return items.filter(Boolean).join(" ");
}

function downloadFile(filename, content, type = "text/plain") {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function exportToDocx(filename, content) {
  // Helper to handle **bold** markdown cleanly
  const parseTextToRuns = (text, isHeader) => {
    if (isHeader) {
      return [new TextRun({ text: text.replace(/\*\*/g, ""), bold: true, size: 22 })];
    }
    const parts = text.split("**");
    return parts.map((part, index) => new TextRun({
      text: part,
      bold: index % 2 === 1,
      size: 22
    }));
  };

  const children = [];
  let tableRows = [];
  let inTable = false;
  let isHeaderRow = false;

  const lines = content.split("\n");
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    
    if (line.startsWith("|") && line.endsWith("|")) {
      if (!inTable) {
        inTable = true;
        isHeaderRow = true;
      }
      
      if (line.match(/^\|(?:[\s-:]+\|)+$/)) {
        continue;
      }

      const cells = line.split("|").slice(1, -1).map(c => c.trim());
      tableRows.push(
        new TableRow({
          tableHeader: isHeaderRow,
          children: cells.map(cellText => new TableCell({
            shading: isHeaderRow ? { fill: "F2F2F2" } : undefined,
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
            children: [
              new Paragraph({
                alignment: isHeaderRow ? AlignmentType.CENTER : AlignmentType.LEFT,
                children: parseTextToRuns(cellText, isHeaderRow),
              })
            ]
          }))
        })
      );
      
      if (isHeaderRow) isHeaderRow = false;

    } else {
      if (inTable) {
        children.push(new Table({
          rows: tableRows,
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 4, color: "7F7F7F" },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: "7F7F7F" },
            left: { style: BorderStyle.SINGLE, size: 4, color: "7F7F7F" },
            right: { style: BorderStyle.SINGLE, size: 4, color: "7F7F7F" },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "D9D9D9" },
            insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "D9D9D9" },
          }
        }));
        children.push(new Paragraph({ text: "", spacing: { after: 200 } })); // Spacer after table
        tableRows = [];
        inTable = false;
      }
      
      let text = line;
      let heading = undefined;
      
      if (line.startsWith("### ")) {
        text = line.replace("### ", "");
        heading = HeadingLevel.HEADING_3;
      } else if (line.startsWith("## ")) {
        text = line.replace("## ", "");
        heading = HeadingLevel.HEADING_2;
      } else if (line.startsWith("# ")) {
        text = line.replace("# ", "");
        heading = HeadingLevel.HEADING_1;
      }

      if (text || heading) {
        children.push(new Paragraph({
          heading: heading,
          spacing: { after: 120, before: heading ? 240 : 0 },
          children: parseTextToRuns(text, false)
        }));
      } else if (line === "") {
        children.push(new Paragraph({ text: "", spacing: { after: 120 } }));
      }
    }
  }
  
  if (inTable) {
    children.push(new Table({
      rows: tableRows,
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: "7F7F7F" },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: "7F7F7F" },
        left: { style: BorderStyle.SINGLE, size: 4, color: "7F7F7F" },
        right: { style: BorderStyle.SINGLE, size: 4, color: "7F7F7F" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "D9D9D9" },
        insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "D9D9D9" },
      }
    }));
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: "Arial", size: 22, color: "262626" },
        },
        heading1: { run: { font: "Arial", size: 32, bold: true, color: "000000" } },
        heading2: { run: { font: "Arial", size: 28, bold: true, color: "000000" } },
        heading3: { run: { font: "Arial", size: 24, bold: true, color: "000000" } },
      },
    },
    sections: [{
      properties: {},
      children: children,
    }],
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function SectionTitle({ icon: Icon, title, subtitle }) {
  return (
    <div className="flex items-start gap-3">
      <div className="rounded-2xl bg-slate-900 p-2 text-white shadow-sm">
        <Icon size={20} />
      </div>
      <div>
        <h2 className="text-xl font-bold text-slate-900">{title}</h2>
        {subtitle && <p className="text-sm text-slate-500">{subtitle}</p>}
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="block">
      <span className="mb-1 block text-sm font-semibold text-slate-700">{label}</span>
      {children}
    </label>
  );
}

function Input(props) {
  return (
    <input
      {...props}
      className={classNames(
        "w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm outline-none transition focus:border-slate-500 focus:ring-2 focus:ring-slate-200",
        props.className
      )}
    />
  );
}

function Textarea(props) {
  return (
    <textarea
      {...props}
      className={classNames(
        "w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm outline-none transition focus:border-slate-500 focus:ring-2 focus:ring-slate-200",
        props.className
      )}
    />
  );
}

function Select(props) {
  return (
    <select
      {...props}
      className={classNames(
        "w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm outline-none transition focus:border-slate-500 focus:ring-2 focus:ring-slate-200",
        props.className
      )}
    >
      {props.children}
    </select>
  );
}

function Button({ children, variant = "dark", className, ...props }) {
  const variants = {
    dark: "bg-slate-900 text-white hover:bg-slate-700",
    light: "bg-white text-slate-800 border border-slate-200 hover:bg-slate-50",
    danger: "bg-rose-50 text-rose-700 border border-rose-100 hover:bg-rose-100",
    gold: "bg-amber-500 text-slate-950 hover:bg-amber-400",
  };
  return (
    <button
      {...props}
      className={classNames(
        "inline-flex items-center justify-center gap-2 rounded-xl px-4 py-2 text-sm font-bold shadow-sm transition disabled:cursor-not-allowed disabled:opacity-50",
        variants[variant],
        className
      )}
    >
      {children}
    </button>
  );
}

export default function AMCOLTOSExamBuilder() {
  const [course, setCourse] = useState({
    institution: "Asian Maritime Technological College (AMCOL)",
    department: "Academic Department",
    courseName: "General Ship Knowledge",
    courseCode: "",
    level: "Vocational / Maritime Program",
    semester: "Semester 1",
    teacher: "",
    totalWeeks: 18,
    hoursPerWeek: 3,
  });

  const [examSettings, setExamSettings] = useState({
    examTitle: "Midterm Examination",
    totalScore: 20,
    totalItems: 40,
  });

  const [objectives, setObjectives] = useState(SAMPLE_OBJECTIVES);
  const [apiKey, setApiKey] = useState("");
  const [aiModel, setAiModel] = useState("gemini-2.5-flash");
  const [aiTOS, setAiTOS] = useState("");
  const [aiExam, setAiExam] = useState("");
  const [isGenerating, setIsGenerating] = useState(false);
  const [activeTab, setActiveTab] = useState("course");
  const [savedNotice, setSavedNotice] = useState("");

  useEffect(() => {
    const saved = localStorage.getItem("amcol_tos_exam_builder_v1");
    if (saved) {
      try {
        const data = JSON.parse(saved);
        if (data.course) setCourse(data.course);
        if (data.examSettings) setExamSettings(data.examSettings);
        if (data.objectives) setObjectives(data.objectives);
        if (data.apiKey) setApiKey(data.apiKey);
        if (data.aiModel) setAiModel(data.aiModel);
        if (data.aiTOS) setAiTOS(data.aiTOS);
        if (data.aiExam) setAiExam(data.aiExam);
      } catch (error) {
        console.warn("Could not load saved data", error);
      }
    }
  }, []);

  function saveData() {
    localStorage.setItem(
      "amcol_tos_exam_builder_v1",
      JSON.stringify({ course, examSettings, objectives, apiKey, aiModel, aiTOS, aiExam })
    );
    setSavedNotice("Saved in this browser.");
    setTimeout(() => setSavedNotice(""), 2200);
  }

  function resetSample() {
    setCourse({
      institution: "Asian Maritime Technological College (AMCOL)",
      department: "Academic Department",
      courseName: "General Ship Knowledge",
      courseCode: "",
      level: "Vocational / Maritime Program",
      semester: "Semester 1",
      teacher: "",
      totalWeeks: 18,
      hoursPerWeek: 3,
    });
    setExamSettings({
      examTitle: "Midterm Examination",
      totalScore: 20,
      totalItems: 40,
    });
    setObjectives(SAMPLE_OBJECTIVES);
    setAiTOS("");
    setAiExam("");
  }

  function addObjective() {
    setObjectives([
      ...objectives,
      { id: crypto.randomUUID(), text: "", hours: 3 },
    ]);
  }

  function updateObjective(id, key, value) {
    setObjectives(objectives.map((o) => (o.id === id ? { ...o, [key]: value } : o)));
  }

  function deleteObjective(id) {
    setObjectives(objectives.filter((o) => o.id !== id));
  }

  async function generateExamWithAI() {
    if (!apiKey) {
      alert("Please enter a Google Gemini API Key first.");
      return;
    }
    setIsGenerating(true);
    setAiTOS("Analyzing objectives and generating TOS blueprint...");
    setAiExam("Waiting to generate exam questions...");
    setActiveTab("results");

    try {
      const prompt = `You are an expert curriculum developer and maritime/vocational instructor.
I will provide the Course Information and Learning Objectives.

Course Information:
- Name: ${course.courseName}
- Level: ${course.level}
- Semester: ${course.semester}
- Total Weeks: ${course.totalWeeks}
- Hours per Week: ${course.hoursPerWeek}

Learning Objectives & Allocated Hours:
${objectives.map((o, i) => `${i + 1}. ${o.text} (${o.hours || 0} hours)`).join("\n")}

Exam Settings:
- Title: ${examSettings.examTitle}
- Total Items: ${examSettings.totalItems}
- Total Score: ${examSettings.totalScore}

Task 1: Generate the Table of Specification (TOS)
Create a Markdown table representing the TOS. Break down the course into logical topics based on the learning objectives.
You MUST format the TOS with the following exact column headers:
| Topic & Lesson Objectives | Hours Taught | % of Exam | Total Items | R1 (Remember) | R2 (Understand) | R3 (Apply) | R4 (Analyze) | R5 (Evaluate) | R6 (Create) | Exact Exam Types | Points Per Item |

For each topic, use the exact allocated teaching hours provided.
For the Bloom's taxonomy columns (R1 to R6), indicate the number of items for that topic that fall under each level.
Ensure the total items and score exactly match the Exam Settings provided above.

Task 2: Generate the Exam Questions
Based strictly on the TOS you just created, write the full examination paper.
Format it cleanly with instructions. Include an Answer Key at the end.

IMPORTANT: You must separate the output for Task 1 and Task 2 exactly with this delimiter text on a new line:
===SPLIT_HERE===
`;

      const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${aiModel}:generateContent?key=${apiKey}`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.7 }
        }),
      });

      const data = await response.json();
      if (data.error) throw new Error(data.error.message);

      const content = data.candidates[0].content.parts[0].text;
      const parts = content.split("===SPLIT_HERE===");
      
      setAiTOS(parts[0] ? parts[0].trim() : "Failed to parse TOS correctly. Please try again.");
      setAiExam(parts[1] ? parts[1].trim() : content); // Fallback to full content if split fails

    } catch (error) {
      console.error(error);
      setAiTOS("Error generating TOS: " + error.message);
      setAiExam("Error generating Exam.");
    } finally {
      setIsGenerating(false);
    }
  }

  const tabClass = (name) =>
    classNames(
      "rounded-xl px-4 py-2 text-sm font-bold transition",
      activeTab === name ? "bg-slate-900 text-white shadow-sm" : "bg-white text-slate-600 hover:bg-slate-100"
    );

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-100 via-white to-amber-50 p-4 text-slate-900 md:p-8">
      <div className="mx-auto max-w-7xl space-y-6">
        <header className="overflow-hidden rounded-3xl bg-slate-950 text-white shadow-xl">
          <div className="grid gap-6 p-6 md:grid-cols-[1.5fr_.8fr] md:p-8">
            <div>
              <div className="mb-4 inline-flex rounded-full border border-amber-300/40 bg-white/10 px-3 py-1 text-xs font-bold uppercase tracking-wider text-amber-200">
                AMCOL Academic Tool (Updated)
              </div>
              <h1 className="text-3xl font-black leading-tight md:text-5xl">
                Auto TOS & Exam Generator
              </h1>
              <p className="mt-3 max-w-2xl text-sm leading-6 text-slate-300 md:text-base">
                Simply input your course information and learning objectives. Our AI Assistant will automatically distribute topics, map Bloom's Taxonomy, create a Table of Specification, and write the full exam paper.
              </p>
            </div>
            <div className="rounded-3xl border border-white/10 bg-white/10 p-5 backdrop-blur">
              <p className="text-sm font-bold text-amber-200">Current Course</p>
              <p className="mt-2 text-2xl font-black">{course.courseName || "Untitled Course"}</p>
              <p className="mt-1 text-sm text-slate-300">{course.courseCode || "No course code yet"}</p>
            </div>
          </div>
        </header>

        <div className="sticky top-3 z-10 rounded-2xl border border-slate-200 bg-white/90 p-2 shadow-sm backdrop-blur print:hidden">
          <div className="flex flex-wrap gap-2">
            <button className={tabClass("course")} onClick={() => setActiveTab("course")}>1. Course Info</button>
            <button className={tabClass("lo")} onClick={() => setActiveTab("lo")}>2. Learning Objectives</button>
            <button className={tabClass("generator")} onClick={() => setActiveTab("generator")}>3. AI Generator</button>
            <button className={tabClass("results")} onClick={() => setActiveTab("results")}>4. Results & Export</button>
            <div className="ml-auto flex flex-wrap gap-2">
              <Button variant="light" onClick={saveData}><Save size={16} /> Save</Button>
              <Button variant="danger" onClick={resetSample}><RefreshCw size={16} /> Reset</Button>
            </div>
          </div>
          {savedNotice && <p className="mt-2 text-xs font-bold text-emerald-600">{savedNotice}</p>}
        </div>

        {activeTab === "course" && (
          <section className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm md:p-6">
            <SectionTitle icon={GraduationCap} title="Course Information" subtitle="Input the official details of the subject/course." />
            <div className="mt-6 grid gap-4 md:grid-cols-2 lg:grid-cols-3">
              <Field label="Institution"><Input value={course.institution} onChange={(e) => setCourse({ ...course, institution: e.target.value })} /></Field>
              <Field label="Department"><Input value={course.department} onChange={(e) => setCourse({ ...course, department: e.target.value })} /></Field>
              <Field label="Teacher / Instructor"><Input value={course.teacher} onChange={(e) => setCourse({ ...course, teacher: e.target.value })} placeholder="Teacher name" /></Field>
              <Field label="Course Name"><Input value={course.courseName} onChange={(e) => setCourse({ ...course, courseName: e.target.value })} /></Field>
              <Field label="Course Code"><Input value={course.courseCode} onChange={(e) => setCourse({ ...course, courseCode: e.target.value })} placeholder="e.g., 20101-2035" /></Field>
              <Field label="Level / Program"><Input value={course.level} onChange={(e) => setCourse({ ...course, level: e.target.value })} /></Field>
              <Field label="Semester"><Input value={course.semester} onChange={(e) => setCourse({ ...course, semester: e.target.value })} /></Field>
              <Field label="Total Weeks"><Input type="number" value={course.totalWeeks} onChange={(e) => setCourse({ ...course, totalWeeks: e.target.value })} /></Field>
              <Field label="Hours per Week"><Input type="number" value={course.hoursPerWeek} onChange={(e) => setCourse({ ...course, hoursPerWeek: e.target.value })} /></Field>
            </div>
          </section>
        )}

        {activeTab === "lo" && (
          <section className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm md:p-6">
            <div className="flex flex-wrap items-start justify-between gap-3">
              <SectionTitle icon={ClipboardList} title="Learning Objectives" subtitle="Add all learning objectives for this course. The AI will map them to Bloom's taxonomy automatically." />
              <Button onClick={addObjective}><Plus size={16} /> Add Objective</Button>
            </div>
            <div className="mt-6 space-y-4">
              {objectives.map((obj, index) => (
                <div key={obj.id} className="flex flex-col gap-3 rounded-2xl border border-slate-200 bg-slate-50 p-4 md:flex-row md:items-center">
                  <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-white font-black text-slate-400 shadow-sm">{index + 1}</div>
                  <Textarea className="flex-1" rows={2} value={obj.text} onChange={(e) => updateObjective(obj.id, "text", e.target.value)} placeholder="Students will be able to..." />
                  <div className="w-full shrink-0 md:w-32">
                    <Input type="number" value={obj.hours} onChange={(e) => updateObjective(obj.id, "hours", e.target.value)} placeholder="Hours" title="Teaching Hours" />
                  </div>
                  <Button variant="danger" className="h-10 shrink-0" onClick={() => deleteObjective(obj.id)}><Trash2 size={16} /></Button>
                </div>
              ))}
            </div>
          </section>
        )}

        {activeTab === "generator" && (
          <section className="space-y-6">
            <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm md:p-6">
              <SectionTitle icon={Settings} title="Target Exam Settings" subtitle="Tell the AI how large you want the exam to be." />
              <div className="mt-6 grid gap-4 md:grid-cols-3">
                <Field label="Exam Title"><Input value={examSettings.examTitle} onChange={(e) => setExamSettings({ ...examSettings, examTitle: e.target.value })} /></Field>
                <Field label="Target Total Items"><Input type="number" value={examSettings.totalItems} onChange={(e) => setExamSettings({ ...examSettings, totalItems: e.target.value })} /></Field>
                <Field label="Target Total Score"><Input type="number" value={examSettings.totalScore} onChange={(e) => setExamSettings({ ...examSettings, totalScore: e.target.value })} /></Field>
              </div>
            </div>

            <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm md:p-6">
              <SectionTitle icon={Sparkles} title="Generate with AI" subtitle="The AI will automatically determine topics, hours, and test formats based on your course info." />
              <div className="mt-6 rounded-2xl border border-amber-200 bg-amber-50 p-4 shadow-sm">
                <Field label="Google Gemini API Key">
                  <div className="flex gap-2">
                    <div className="relative flex-1">
                      <Key className="absolute left-3 top-2.5 text-slate-400" size={16} />
                      <Input
                        type="password"
                        placeholder="AIza..."
                        className="pl-9"
                        value={apiKey}
                        onChange={(e) => setApiKey(e.target.value)}
                      />
                    </div>
                    <Button onClick={saveData} variant="light">Save Key</Button>
                  </div>
                </Field>
                <div className="mt-4">
                  <Field label="AI Model">
                    <Select value={aiModel} onChange={(e) => setAiModel(e.target.value)} className="mt-2">
                      <option value="gemini-2.5-pro">Gemini 2.5 Pro</option>
                      <option value="gemini-2.5-flash">Gemini 2.5 Flash</option>
                      <option value="gemini-2.0-pro-exp-0205">Gemini 2.0 Pro Experimental</option>
                      <option value="gemini-2.0-flash">Gemini 2.0 Flash</option>
                      <option value="gemini-1.5-pro">Gemini 1.5 Pro</option>
                      <option value="gemini-1.5-flash">Gemini 1.5 Flash</option>
                    </Select>
                  </Field>
                </div>
                <p className="mt-4 text-xs text-amber-700">
                  Your API key is stored locally in your browser and sent directly to Google. We do not store it on any external servers.
                </p>
              </div>

              <div className="mt-6">
                <Button onClick={generateExamWithAI} disabled={isGenerating} variant="gold" className="w-full py-3 text-base">
                  <Sparkles size={20} />
                  {isGenerating ? "Generating Course Materials..." : "Generate Full TOS & Exam Paper"}
                </Button>
              </div>
            </div>
          </section>
        )}

        {activeTab === "results" && (
          <section className="space-y-6">
            {aiTOS && (
              <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm md:p-6">
                <div className="flex items-center justify-between">
                  <SectionTitle icon={ClipboardList} title="Table of Specification" />
                  <Button variant="light" onClick={() => exportToDocx(`Generated_TOS_${course.courseCode || 'AI'}.docx`, aiTOS)}><Download size={16} /> Export to Word</Button>
                </div>
                <div className="mt-4 overflow-x-auto rounded-2xl bg-slate-50 p-4 text-sm text-slate-800 shadow-inner">
                  <pre className="whitespace-pre-wrap font-sans">{aiTOS}</pre>
                </div>
              </div>
            )}
            
            {aiExam && (
              <div className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm md:p-6">
                <div className="flex items-center justify-between">
                  <SectionTitle icon={FileText} title="Generated Examination" />
                  <Button variant="light" onClick={() => exportToDocx(`Generated_Exam_${course.courseCode || 'AI'}.docx`, aiExam)}><Download size={16} /> Export to Word</Button>
                </div>
                <div className="mt-4 max-h-[800px] overflow-y-auto rounded-2xl bg-slate-50 p-4 text-sm leading-relaxed text-slate-800 shadow-inner">
                  <pre className="whitespace-pre-wrap font-sans">{aiExam}</pre>
                </div>
              </div>
            )}
          </section>
        )}

        <footer className="pb-6 text-center text-xs text-slate-400 print:hidden">
          AMCOL TOS & Examination Builder — data is stored locally in the browser unless exported.
        </footer>
      </div>
    </div>
  );
}
