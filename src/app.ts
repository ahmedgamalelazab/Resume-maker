// Generate a CV
import * as fs from "fs";

import { AlignmentType, Document, HeadingLevel, Packer, Paragraph, Tab, TabStopPosition, TabStopType, TextRun } from "docx";

// tslint:disable:no-shadowed-variable

const PHONE_NUMBER = "201032122442";
const PROFILE_URL = "https://www.linkedin.com/in/ahmed-el-azab-774704237/";
const EMAIL = "agemy844@gmail.com";


interface Experience {
    readonly isCurrent: boolean;
    readonly summary: string[];
    readonly title: string;
    readonly startDate: {
        readonly month: number;
        readonly year: number;
    };
    readonly endDate?: {
        readonly month: number;
        readonly year: number;
    };
    readonly company: {
        readonly name: string;
    };
}

interface Education {
    readonly degree: string;
    readonly fieldOfStudy: string;
    readonly notes: string[];
    readonly schoolName: string;
    readonly startDate: {
        readonly year: number;
    };
    readonly endDate: {
        readonly year: number;
    };
}

interface Skill {
    readonly name: string;
}

interface Achievement {
    readonly issuer: string;
    readonly name: string;
}

const experiences: Experience[] = [
    {
        isCurrent: true,
        summary: [
            "Demonstrated expertise in integrating Asterisk* with VOIP systems.",
            "Proficient in using Node.js with TypeScript for web development.",
            "Applied CQRS architecture pattern and maintained a clean codebase for backend development.",
            "Utilized messaging tools like REDIS for efficient data handling.",
            "Managed NoSQL databases, specifically MongoDB.",
            "Leveraged AWS and GCP services for various aspects of software development.",
            "Worked with frontend frameworks such as AngularJS and Angular."
        ],
        title: "Software engineer",
        startDate: {
            month: 10,
            year: 2022,
        },
        company: {
            name: "AppoutITS",
        },
    },
    {
        isCurrent: false,
        summary:[
            "Successfully integrated CRM Dynamics 365 into the company's software ecosystem.",
            "Developed and maintained .NET Core Web APIs for backend services.",
            "Utilized SQL, specifically MSSQL, for efficient data storage and retrieval.",
            "Contributed to open-source .NET E-commerce projects, particularly nopCommerce."
        ],
        title: "Software Developer",
        endDate: {
            month: 9,
            year: 2022,
        },
        startDate: {
            month: 5,
            year: 2022,
        },
        company: {
            name: "Excellent Protection",
        },
    },
];

const education: Education[] = [
    {
        degree: "Diploma of Education, Computer Engineering",
        fieldOfStudy: "Computer Science",
        notes: [
            "The program is offered as a full-fledged scholarship for selected Egyptian university graduates within five calendar years of their graduation. It acts as a catalyst to bridge the gap between the talent supply skills and the domestic, regional, and international market demands",
            "Collaborating with a rich network of industry partners, more than 32 various tracks; in which candidates join after a rigorous screening process for 8 weeks and 10% acceptance rate; were designed and updated in terms of target job profiles and ICT market competencies and serious, committed and passionate learners."
        ],
        schoolName: "Information Technology Institute ( ITI ) 9 month",
        startDate: {
            year: 2021,
        },
        endDate: {
            year: 2022,
        },
    },
    {
        degree: "Bachelor's degree, Engineering",
        fieldOfStudy: "Civil engineering",
        notes: [
            "Started programming at the 3rd year",
            "Learned Excel VBA programming language to automate civil engineering construction sheets",
            "Learned Fortran to write AutoCAD Lisp scripts to automate repeated drawing tasks"
        ],
        schoolName: "Behira High Institute of Engineering and Technology",
        startDate: {
            year: 2014,
        },
        endDate: {
            year: 2019,
        },
    }
    
];

const skills: Skill[] = [
    {
        name: "Angular",
    },
    {
        name: "TypeScript",
    },
    {
        name: "JavaScript",
    },
    {
        name: "NodeJS",
    },
    {
        name: "Redis",
    },
    {
        name: "Git",
    },
    {
        name: "AWS - GCP",
    },
    {
        name: "VOIP (Asterisk*)",
    },
    {
        name: "SQL( Postgres - mssql )",
    },
    {
        name: "Nosql (Mongodb)",
    },
    {
        name: "OOP",
    },
];

const achievements: Achievement[] = [
    {
        issuer: "Information Technology Institute (ITI)",
        name: "ITI cross platform mobile developing intake 42 ( ITI ) 9 month scholar ship",
    },
    {
        issuer: "Udacity ( Udacity Nanodegree Graduation Certificate )",
        name: "Angular development cross skilling",
    },
];

class DocumentCreator {
    // tslint:disable-next-line: typedef
    public create([experiences, educations, skills, achievements]: [Experience[], Education[], Skill[], Achievement[]]): Document {
        const document = new Document({
            sections: [
                {
                    children: [
                        new Paragraph({
                            text: "Ahmed Gamal Mohamed Elazab",
                            heading: HeadingLevel.TITLE,
                            alignment: AlignmentType.CENTER
                        }),
                        this.createContactInfo(PHONE_NUMBER, PROFILE_URL, EMAIL),
                        this.createHeading("Education"),
                        ...educations
                            .map((education) => {
                                const arr: Paragraph[] = [];
                                arr.push(
                                    this.createInstitutionHeader(
                                        education.schoolName,
                                        `${education.startDate.year} - ${education.endDate.year}`,
                                    ),
                                );
                                arr.push(this.createRoleText(`${education.fieldOfStudy} - ${education.degree}`));
                            
                                education.notes.forEach((bulletPoint) => {
                                    arr.push(this.createBullet(bulletPoint));
                                });

                                return arr;
                            })
                            .reduce((prev, curr) => prev.concat(curr), []),
                        this.createHeading("Experience"),
                        ...experiences
                            .map((position) => {
                                const arr: Paragraph[] = [];

                                arr.push(
                                    this.createInstitutionHeader(
                                        position.company.name,
                                        this.createPositionDateText(position.startDate, position.endDate, position.isCurrent),
                                    ),
                                );
                                arr.push(this.createRoleText(position.title));
                            

                                position.summary.forEach((bulletPoint) => {
                                    arr.push(this.createBullet(bulletPoint));
                                });

                                return arr;
                            })
                            .reduce((prev, curr) => prev.concat(curr), []),
                        this.createHeading("Skills, Achievements and Interests"),
                        this.createSubHeading("Skills"),
                        this.createSkillList(skills),
                        this.createSubHeading("Achievements"),
                        ...this.createAchievementsList(achievements),
                        this.createSubHeading("Interests"),
                        this.createInterests("Programming, Technology, Music Production, Web Design, 3D Modeling, Dancing."),
                        this.createHeading("References"),
                        new Paragraph(
                            "Eng Abdullah mohamed ITI instructor, he was my instructor and my leader at ITI, you can reach him via phone: 201021905757",
                        ),
                        new Paragraph(
                            "Eng Nasr kassem ITI instructor, he was my instructor and our session lead, you can reach him via phone: 201090290520",
                        ),
                        // new Paragraph("More references upon request"),
                        // new Paragraph({
                        //     text: "This CV was generated in real-time based on my Linked-In profile from my personal website www.dolan.bio.",
                        //     alignment: AlignmentType.CENTER,
                        // }),
                    ],
                },
            ],
        });

        return document;
    }

    

    public createContactInfo(phoneNumber: string, profileUrl: string, email: string): Paragraph {
        return new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
                new TextRun(`Mobile: ${phoneNumber} | LinkedIn: ${profileUrl} | Email: ${email}`),
                new TextRun({
                    text: "Address: Egypt Giza (Sheekh Zaid City)",
                    break: 1,
                }),
            ],
        });
    }

    public createHeading(text: string): Paragraph {
        return new Paragraph({
            text: text,
            heading: HeadingLevel.HEADING_1,
            thematicBreak: true,
        });
    }

    public createSubHeading(text: string): Paragraph {
        return new Paragraph({
            text: text,
            heading: HeadingLevel.HEADING_2,
        });
    }

    public createInstitutionHeader(institutionName: string, dateText: string): Paragraph {
        return new Paragraph({
            tabStops: [
                {
                    type: TabStopType.RIGHT,
                    position: TabStopPosition.MAX,
                },
            ],
            children: [
                new TextRun({
                    text: institutionName,
                    bold: true,
                }),
                new TextRun({
                    children: [new Tab(), dateText],
                    bold: true,
                }),
            ],
        });
    }

    public createRoleText(roleText: string): Paragraph {
        return new Paragraph({
            children: [
                new TextRun({
                    text: roleText,
                    italics: true,
                }),
            ],
        });
    }

    public createBullet(text: string): Paragraph {
        return new Paragraph({
            text: text,
            bullet: {
                level: 0,
            },
        });
    }

    // tslint:disable-next-line:no-any
    public createSkillList(skills: any[]): Paragraph {
        return new Paragraph({
            children: [new TextRun(skills.map((skill) => skill.name).join(", ") + ".")],
        });
    }

    // tslint:disable-next-line:no-any
    public createAchievementsList(achievements: any[]): Paragraph[] {
        return achievements.map(
            (achievement) =>
                new Paragraph({
                    text: achievement.name,
                    bullet: {
                        level: 0,
                    },
                }),
        );
    }

    public createInterests(interests: string): Paragraph {
        return new Paragraph({
            children: [new TextRun(interests)],
        });
    }

    public splitParagraphIntoBullets(text: string): string[] {
        return text.split("\n\n");
    }

    // tslint:disable-next-line:no-any
    public createPositionDateText(startDate: any, endDate: any, isCurrent: boolean): string {
        const startDateText = this.getMonthFromInt(startDate.month) + ". " + startDate.year;
        const endDateText = isCurrent ? "Present" : `${this.getMonthFromInt(endDate.month)}. ${endDate.year}`;

        return `${startDateText} - ${endDateText}`;
    }

    public getMonthFromInt(value: number): string {
        switch (value) {
            case 1:
                return "Jan";
            case 2:
                return "Feb";
            case 3:
                return "Mar";
            case 4:
                return "Apr";
            case 5:
                return "May";
            case 6:
                return "Jun";
            case 7:
                return "Jul";
            case 8:
                return "Aug";
            case 9:
                return "Sept";
            case 10:
                return "Oct";
            case 11:
                return "Nov";
            case 12:
                return "Dec";
            default:
                return "N/A";
        }
    }
}

const documentCreator = new DocumentCreator();

const doc = documentCreator.create([experiences, education, skills, achievements]);

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});