import shinyswatch
from shiny import App, Inputs, Outputs, Session, render, ui, reactive
import pandas
from docxtpl import DocxTemplate
# from docxtpl import DocxTemplate
import numpy as np
import asyncio
import io
from datetime import date
from pathlib import Path
import win32com.client as win32 
import os,sys
from citeproc import CitationStylesStyle, CitationStylesBibliography
from citeproc.source.json import CiteProcJSON
from citeproc.source.bibtex import BibTeX
# import win32com.client as win32 
# from citeproc import CitationStylesStyle, CitationStylesBibliography
# from citeproc.source.json import CiteProcJSON
# from citeproc.source.bibtex import BibTeX
import openai
import pickle

# pip install citeproc-py
# https://pypi.org/project/citeproc-py/



#GPT API Setting
API_KEY = "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"
openai.api_key = API_KEY
model_id = 'gpt-3.5-turbo'

os.chdir(sys.path[0])

# A card component wrapper. (for download)
def ui_card(title, *args):
    return (
        ui.div(
            {"class": "card mb-4"},
            ui.div(title, class_="card-header"),
            ui.div({"class": "card-body"}, *args),
        ),
    )


app_ui = ui.page_navbar(
    shinyswatch.theme.superhero(),
    ui.nav( 'The Syllabus Build Guide',
        ui.panel_main(
    ui.nav( '',

            ui.navset_tab(
                ui.nav(
                    "Info",       ui.HTML('<br>'),
                        ui.input_text_area('programname','Program Name','',placeholder=' [Program Name]',width='100%',height='10px'),
                        ui.input_text_area('classnameandnumber','Class Name and Number','',placeholder="[Course Title and Number]",width='100%',height='10px'),
                        ui.input_text_area('classtime','Class Time','',placeholder="[Scheduled Meeting Times]",width='100%',height='10px'),
                        ui.input_text_area('credits','Credits','',placeholder='[Number of credits]',width='100%',height='10px'),
                        ui.input_text_area('coursetype','Course Type','',placeholder='[Core course or Elective]',width='100%',height='10px'),                         
                        ui.input_text_area('instructor','Instructor','',placeholder='[Name, title, email address and phone number]',width='100%',height='10px'),
                        ui.input_text_area('officehours','Office Hours','',placeholder='[SPS Policy: Must state date, time and location; may also indicate by appointment]',width='100%',height='10px'),
                        ui.input_text_area('responsepolicy','Response Policy:','',placeholder='[Include a brief statement about your preferred means of communication and when students should expect a response from you. Will you be available 24/7 or during the workweek only? Will you generally respond within 12 or 24 hours?]',width='100%',height='64px'),
                        ui.input_text_area('TA','Facilitator/Teaching Assistant,','',placeholder='[Name, title, email address and phone number]',width='100%',height='10px'),
                        ui.input_text_area('TAofficehours','TA Office Hours','',placeholder='[SPS policy: Must state date, time and location; may also indicate by appointment]',width='100%',height='10px'),
                        ui.input_text_area('TAresponsepolicy','Response Policy:','',placeholder='[Include a brief statement about your preferred means of communication and when students should expect a response from you. Will you be available 24/7 or during the workweek only? Will you generally respond within 12 or 24 hours?]',width='100%',height='64px'),
                        ),
                ui.nav(
                    "Overview",       ui.HTML('<br>'),
                        ui.input_text_area('courseoverview1','First Paragraph','',placeholder='''(a)	Provide a stimulating and descriptive overview of the course. Be sure to include:     
                                           i. the course’s main topics    
                                           ii. for whom the course is designed (e.g., for everyone in the program or primarily for those pursuing a special track)''',width='100%',height='200px'),
                        ui.input_text_area('courseoverview2','Second Paragraph','',placeholder='''(b)	Identify the larger programmatic goals that the course serves. Include:
                                            i. how the course relates to the primary concepts and principles of the discipline
                                            ii. how the course fits in with the program curriculum
                                            ''',width='100%',height='200px'),
                        ui.input_text_area('courseoverview3','Third Paragraph','',placeholder='''(c)	Course logistics. Indicate:
                                            i. whether the course is a required core course or an elective
                                            ii. whether or not it will be open, space permitting, to cross-registrants from other fields and/or Columbia University programs; if so which ones
                                            iii. whether specific competencies or prerequisite knowledge or course work in the discipline are required
                                            iv. Course Modality (Describe delivery modality: e.g., online, on-campus, hybrid/Hy-flex
                                            v. Duration. Describe whether the course is: Full semester  Block Week, Partial semester, Residencies; or Other format]
                                            ''',width='100%',height='200px'),
                                     
                ),
                ui.nav(
                    "Objectives ",       ui.HTML('<br>'),
                        ui.output_text_verbatim('textout_LearningOutcomes'),
                        ui.input_text_area('l1','L1',width='100%'),
                        ui.input_text_area('l2','L2',width='100%'),
                        ui.input_text_area('l3','L3',width='100%'), 
                        ui.input_text_area('l4','L4',width='100%'), 
                        ui.input_text_area('l5','L5',width='100%'), 
                        ui.input_text_area('l6','L6',width='100%'),   
                          
                ),
                ui.nav(
                     "Readings", ui.HTML('<br>'),
                     {'id' :'readings'},
                          
                        ui.input_select("citation", "Select Your citation Style", {"APA": "APA", "MLA": "MLA", "Chicago": "Chicago"}),
                        ui.input_text_area('citationinfo','Citation information(Replace the information inside[] )','',placeholder='This is a [Book/Website/Video], the Author(s) is [ NULL ], Title of the book/webpage/article is [ NULL ], Year of publication/Date accessed is [ NULL ], Publisher/Title of the journal is [ NULL ], Page numbers is [ NULL ], Other information: [ NULL ]',width='100%'),
                        ui.input_action_button('action_send','Send'),
                        ui.output_text('citecomplete', 'Cite result'),
                        ui.input_text_area('books','Books (Copy the citation into this box)','',placeholder='''[Identify required and recommended readings for the course. Required readings should include a balance of graduate-level practitioner texts and primary academic sources (scholarly articles from peer-reviewed journals in the discipline). Texts have sufficient breadth, depth, and currency for the student to learn the subject at a Master's level and achieve the stated course learning objectives. 
                                            Provide full citations (author, publisher, publication year, etc.), using a recognized citation format, such as MLA, APA or Chicago Style format, after consultation with your academic director. Include page numbers, page counts, and media listening/viewing times so that students can assess the reading workload. Indicate to students where they may find the materials (e.g., Canvas folders, library, purchase from vendor, etc.). Include web links where relevant.
                                            ''',width='100%',height='200px'),
                        ui.input_text_area('others','Other Required Readings (Copy the citation into this box)','',placeholder= 'Other Required Readings (available through Canvas course site or web link)',width='100%',height='200px'),
                        ui.input_text_area('webandvideo','Websites and Videos (Copy the citation into this box)',width='100%',height='200px'),
                         

                ),
                ui.nav(
                     "Assignments",  ui.HTML('<br>'),
                    {'id' :'assignment'},             
                        ui.output_text_verbatim('textout_Assignments'),                                  
                        ui.input_text_area('writeassignment','Written assignments',width='100%',height='100%'),
                        ui.input_text_area('present','Presentations and performances',width='100%'),
                        ui.input_text_area('exams','Exams',width='100%'),
                        ui.input_text_area('practice','Practice',width='100%'),
                        ui.input_text_area('onlineinteraction','Online Interaction ',width='100%'),
                        ui.input_text_area('participation','Participation',width='100%'),
                        ui.input_text_area('otherassignment','Others',width='100%'),
                         
                ),
                ui.nav(
                     "Grading", ui.HTML('<br>'),
                     ui.div({'id' :'main-content'},
                     ui.input_action_button("btn", "Insert assignment"),
                     
                     )
                ),
                ui.nav(
                     "Schedule",ui.HTML('<br>'),
                     ui.div({'id' :'coursecalendar'},
                     ui.input_action_button("btn2", "Insert a Week"),
                     
                     )                                                
                ),
                ui.nav(
                    "Policies",       ui.HTML('<br>'),
                        ui.input_select('participantion', 'Participation and Attendance', {'You are expected to complete all assigned readings, attend all class sessions, and engage with others in online discussions. Your participation will require that you answer questions, defend your point of view, and challenge the point of view of others. If you need to miss a class for any reason, please discuss the absence with me in advance.' : 'You are expected to complete all assigned readings, attend all class sessions, and engage with others in online discussions. Your participation will require that you answer questions, defend your point of view, and challenge the point of view of others. If you need to miss a class for any reason, please discuss the absence with me in advance. ','I expect you to come to class on time and thoroughly prepared. I will keep track of attendance and look forward to an interesting, lively and confidential discussion. If you miss an experience in class, you miss an important learning moment and the class misses your contribution. More than one absence will affect your grade' : 'I expect you to come to class on time and thoroughly prepared. I will keep track of attendance and look forward to an interesting, lively and confidential discussion. If you miss an experience in class, you miss an important learning moment and the class misses your contribution. More than one absence will affect your grade'},width='100%'),
                        ui.input_select('latework', 'Late work', {'There will be no credit granted to any written assignment that is not submitted on the due date noted in the course syllabus without advance notice and permission from the instructor.':'There will be no credit granted to any written assignment that is not submitted on the due date noted in the course syllabus without advance notice and permission from the instructor.','Work that is not submitted on the due date noted in the course syllabus without advance notice and permission from the instructor will be graded down 1/3 of a grade for every day it is late (e.g., from a B+ to a B).':'Work that is not submitted on the due date noted in the course syllabus without advance notice and permission from the instructor will be graded down 1/3 of a grade for every day it is late (e.g., from a B+ to a B).'},width='100%'),
                        ui.input_text_area('citationpolicy','Citation & Submission','''[All written assignments must use standard citation format (e.g., MLA, APA, Chicago), cite sources, and be submitted to the course website (not via email).]''',width='100%'),     
                ),
                ui.nav(
                    "Resources",   ui.HTML('<br>'),    
                        ui.input_select('onlineclass', 'Does this course will use online platform?',{'yes':'Yes','no':'No'}),
                        ui.panel_conditional(
                        "input.onlineclass === 'yes' ", ui.input_text_area("online", "Online platforms policy(No need change)",''' Netiquette
                        Online sessions in this course will be offered through Zoom, accessible through Canvas.  A reliable Internet connection and functioning webcam and microphone are required. It is your responsibility to resolve any known technical issues prior to class. Your webcam should remain turned on for the duration of each class, and you should expect to be present the entire time. Avoid distractions and maintain professional etiquette. 
                        Please note: Instructors may use Canvas or Zoom analytics in evaluating your online participation.
                        More guidance can be found at: https://jolt.merlot.org/vol6no1/mintu-wimsatt_0310.htm
                        Netiquette is a way of defining professionalism for collaborations and communication that take place in online environments. Here are some Student Guidelines for this class:
                        ●	Avoid using offensive language or language that is not appropriate for a professional setting.
                        ●	Do not criticize or mock someone’s abilities or skills.
                        ●	Communicate in a way that is clear, accurate and easy for others to understand.
                        ●	Balance collegiality with academic honesty.
                        ●	Keep an open-mind and be willing to express your opinion.
                        ●	Reflect on your statements and how they might impact others.
                        ●	Do not hesitate to ask for feedback.
                        ●	When in doubt, always check with your instructor for clarification.
                        ''',width='100%',height='100%')
                        ),
                        

                ),
                ui.nav(
                    'i/o',ui.HTML('<br>'),
                    ui_card(
                        ui.download_button("download0", "Download your docx version class syllabus!"),
                    ),
                    ui_card(
                        ui.download_button("download1", "Download your PDF version class syllabus!"),
                    ),
                    ui.output_text_verbatim('down'), 
                )
            ),
            
),
title="Columbia University School of Professional Studies: Syllabus Builder",
)






def server(input: Inputs, output: Outputs, session: Session):

    @output
    @render.text
    def textout_Assignments():
         return '''[Describe here and enumerate the major graduate-level assignments of the course. 
    These descriptions should be high-level to afford flexibility in an approved syllabus. 
    Detailed descriptions should be contained in the Canvas course site.
    Assignments include all required work to be produced by students and evaluated by the instructor, including: 
    ●	Written assignments (e.g., case analyses, research projects, project plans, reaction papers, essays, designs, op-eds, etc.)
    ●	Presentations and performances (e.g., role-playing, strategic interactions, leading discussions, client meetings, etc.)
    ●	Exams (e.g., tests, mid-terms, in-class assessments, final exams, etc.)
    ●	Practice (e.g., drafts of required written, designed, or performed work, practice sets, etc.)
    ●	Online Interaction (synchronous or asynchronous, e.g., discussions, posts, threads, chats, etc.) 
    ●	Participation (assign no more than 15% of the final grade to participation. Consult with your Academic Director as to program-specific participation grading cap) 
    ●	Other

Include statements regarding 
    1) how assignments help students achieve the stated learning objectives, build skills toward culminating project or exam, and develop competencies that align with the field/discipline, 
    2) pitch and degree of difficulty for the intended audience, 
    3) how you will measure students’ progress toward the course goals (formative assessment),  
    4) specific criteria you will use to evaluate students’ work, and 5) how and when you will provide feedback. Each of these assignments should indicate the learning objectives stated above (L1, L2, etc.). Indicate the grade weight for each assignment and whether the grade is assigned to the individual or to the group/team. Where applicable, please refer students to the Canvas course site for further specificity on assignments.]
'''

    @output
    @render.text
    def textout_LearningOutcomes():
        return '''[Graduate-level learning objectives encompass learning outcomes that require higher-level functioning, critical analysis, and application to professional fields. 
        Such learning objectives will include observable and actionable verbs such as analyze, critique, design, apply, evaluate, etc. Most SPS courses define 4-6 objectives. 
        Consult a one-page primer from Columbia’s Mailman School. See an example of an SPS graduate course syllabus here. 
        SPS Instructional Design team members can also help you with writing objectives aligned with program goals. 
        Please contact the Senior Director of Instructional Design and Curriculum Support, Ariel Fleurimond, af2830@columbia.edu.
These course-level learning objectives should align with programmatic objectives and be: 
•	observable and measurable
•	designed for the level and purpose of the course
•	be focused on the what the learner will do (not what the instructor will teach)
•	labeled L1, L2, etc. and linked to assignments and activities in the appropriate section.]
'''
    
    #Create reactive value to keep btn
    contet = reactive.Value()
    contet2 = reactive.Value()
    
    
    # GPT conversation function
    def ChatGPT_conversation(conversation):
        response = openai.ChatCompletion.create(
            model=model_id,
            messages=conversation
        )
        conversation.append({'role': response.choices[0].message.role, 'content': response.choices[0].message.content})
        return conversation

    
    # Read the Syllabus Template
    # doc = DocxTemplate("SPS Syllabus Template.docx")

    
    #ChatGPT citation
    @reactive.event(input.action_send)
    def citepush():
        conversation = []
        prompt = f'Cite this book by {input.citation()} style format for me: {input.citationinfo()}'
        conversation.append({'role': 'user', 'content': prompt})
        conversation = ChatGPT_conversation(conversation)
        response = ('{0}: {1}\n'.format(conversation[-1]['role'].strip(), conversation[-1]['content'].strip()))
        return response
    
    #citation textout
    @output
    @render.text
    def citecomplete():
        return citepush()
    

    
    #reactive dynamically add input (Grade part) 
    @reactive.Effect
    @reactive.event(input.btn)
    def Grade2():
        ptn = input.btn()        
        if ptn > 0:
            assignment = ui.input_text(f'name{input.btn()}', f"Assignment{input.btn()}name", value=input.btn())
            text = ui.input_text(f'{input.btn()}', f"Assignment{input.btn()}weight", value=input.btn())
            select = ui.input_select(f'type{input.btn()}','type',{"Individual grade": "individual grade", "Group grade": "Group grade"})
            ui.insert_ui(
                ui.div({"id": f"inserted-assignment{input.btn()}"}, assignment),
                selector="#main-content",
                where="beforeEnd",
            )
            ui.insert_ui(
                ui.div({"id": f"inserted-text{input.btn()}"}, text),
                selector="#main-content",
                where="beforeEnd",
            )
            ui.insert_ui(
                ui.div({"id": f"inserted-select{input.btn()}"}, select),
                selector="#main-content",
                where="beforeEnd",
            )
        qqq = str(ptn)
        contet.set(qqq)

    
    #(Grade part)
    def Grade():
        ptn = contet()  
        weight = []
        content = []
        for i in range(int(ptn)): 
            input_id = f'{i+1}'
            assignment = input[f'name{input_id}']() 
            weight = input[input_id]()
            type = input[f'type{input_id}']()
            entry = {'assignment': f'{assignment}', 'weight': f'{weight}%', 'type': f'{type}'}
            content.append(entry)
        return content

    
    #reactive dynamically add input (course calender part)
    @reactive.Effect
    @reactive.event(input.btn2)
    def coursecalender():
        ptn2 = input.btn2()        
        if ptn2 > 0:
            text = ui.input_text(f'date{input.btn2()}', f"Week{input.btn2()}Date", value=input.btn2())
            text1 = ui.input_text(f'topic{input.btn2()}', f"Week{input.btn2()}topic", value='''Course introductions Foundations of … ''',width='100%')
            text2 = ui.input_text(f'reading{input.btn2()}', f"Week{input.btn2()}reading", value='''Title/author Chapters 1–2, pp 105-135(30 pages) Articles x,y,z, pp 24-44 (20 pages) ''',width='100%')
            text3 = ui.input_text(f'assignments{input.btn2()}', f"Week{input.btn2()}assignments", value='''Statement of purpose due 9/15 ''',width='100%')
            ui.insert_ui(
                ui.div({"id": f"inserted-text{input.btn2()}"}, text),
                selector="#coursecalendar",
                where="beforeEnd",
            )
            ui.insert_ui(
                ui.div({"id": f"inserted-text1{input.btn2()}"}, text1),
                selector="#coursecalendar",
                where="beforeEnd",
            )
            ui.insert_ui(
                ui.div({"id": f"inserted-text2{input.btn2()}"}, text2),
                selector="#coursecalendar",
                where="beforeEnd",
            )
            ui.insert_ui(
                ui.div({"id": f"inserted-text3{input.btn2()}"}, text3),
                selector="#coursecalendar",
                where="beforeEnd",
            )
        qqq2 = str(ptn2)
        contet2.set(qqq2)
    
    #(course calender part)
    def Week():
        ptn2 = contet2()
        date = []
        topic = []
        reading = []
        assignments = []
        content2 = []
        for i in range(int(ptn2)): 
            input_id2 = f'{i+1}'
            date = input[f'date{input_id2}']()
            topic = input[f'topic{input_id2}']()
            reading = input[f'reading{input_id2}']()
            assignments = input[f'assignments{input_id2}']()
            entry2 = {'date' : f'Week{i+1}{date}', 'topic': f'{topic}', 'reading': f'{reading}%', 'assignments': f'{assignments}'}
            content2.append(entry2)
        return content2

    
         
    
    
    #render content to the doc
    @output
    @render.text
    def down():
        
        #Create dynamcic input and bullets for learning objectives
        bullets = [
                        input.l1(),
                        input.l2(),
                        input.l3(),
                        input.l4(),
                    ]
        if input.l5() != '':
                bullets += (input.l5(),)
        if input.l6() != '':
                bullets += (input.l6(),)

        particapation = input.participation()
        if input.otherassignment() != '':
                particapation += f'\n\n7. Others\n{input.otherassignment()}'

        if input.onlineclass() == 'no':
            online = ''
        else:
            online = input.online()

        grade = Grade()
        week = Week()
        context = { 
            'programname' : input.programname(),
            'classnameandnumber' : input.classnameandnumber(),
            'classtime' : input.classtime(),
            'credits' : input.credits(),
            'coursetype' : input.coursetype(),
            'instructor' : input.instructor(),
            'officehours' : input.officehours(),
            'responsepolicy' : input.responsepolicy(),
            'TA' : input.TA(),
            'TAofficehours' : input.TAofficehours(),
            'TAresponsepolicy' : input.TAresponsepolicy(),
            'courseoverview1' : input.courseoverview1(),
            'courseoverview2' : input.courseoverview2(),
            'courseoverview3' : input.courseoverview3(),
            'bullets' : bullets,
            'books' : input.books(),
            'others' : input.others(),
            'webandvideo' : input.webandvideo(),
            'writeassignment' : input.writeassignment(),
            'present' : input.present(),
            'exams' : input.exams(),
            'practice' : input.practice(),
            'onlineinteraction' : input.onlineinteraction(),
            'participation' : particapation,
            'tbl_contents': grade,
            'week_table': week,
            'participation': input.participation(),
            'latework': input.latework(),
            'citationpolicy': input.citationpolicy(),
            'online': online       
            } 

        doc.render(context)
        doc.save("Complete.docx")

        print('Already send all the information, please wait for processing!')

    
    #converting doc to pdf
    def convert_to_pdf(doc):
        word = win32.DispatchEx("Word.Application")
        new_name = doc.replace(".docx", r".pdf")
        worddoc = word.Documents.Open(doc)
        worddoc.SaveAs(new_name, FileFormat = 17)
        worddoc.Close()
        return None

    #doc download
    @session.download()
    def download0():
        # This is the simplest case. The implementation simply returns the path to a
        # file on disk.
        path = Path(__file__).parent / "Complete.docx"
        return str(path)

    #pdf download
    @session.download()
    def download1():
        # This is the simplest case. The implementation simply returns the path to a
        # file on disk.
        path_to_word_document = os.path.join(os.getcwd(), f'Complete.docx')
        convert_to_pdf(path_to_word_document)
        path = Path(__file__).parent / "Complete.pdf"
        return str(path)


app = App(app_ui, server)