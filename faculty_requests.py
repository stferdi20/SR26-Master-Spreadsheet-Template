"""SR25-style tailored reporting requests for faculties, keyed by S2030 'Qual' worksheet row.

Each entry: (request, note). `note` states what SST already holds or will collect elsewhere
(SR26 approach 4.2: say why the recipient is asked and what information is already held).
The original Qual-sheet wording is kept in the master's 'Request wording log' tab.
"""

TAILORED = {
    8: ("Please provide one or two examples from 2026 of innovative approaches your faculty has used to embed sustainability "
        "in teaching and learning (e.g. new or redesigned subjects, curriculum design, pedagogies, assessment approaches or "
        "learning resources). Include links to subject pages or resources where available.",
        "A full list of subjects is not required."),
    9: ("Has your faculty piloted or trialled any approach in 2026 to understand or build students' confidence and capability "
        "to act on sustainability (sustainability self-efficacy)? If yes, briefly describe the approach, participating cohort, "
        "what was learned and next steps. If not, please respond 'none in 2026' – this is a new indicator and a nil response is useful.",
        "We are aware of the Sustainability Strengths Compass Living Lab project (CLLAP) and will collect it separately."),
    10: ("Please list any sustainability-focused experiential or co-curricular learning opportunities your faculty offered or "
         "supported in 2026 (e.g. industry or community projects, fieldwork, internships, student challenges), with approximate "
         "student participation where known and a one-line description of outcomes.",
         "We will collect Wattle Fellowship and Melbourne Plus information centrally – please only include these if your faculty ran a specific activity."),
    11: ("Please nominate up to two alumni from your faculty whose careers or community work demonstrate how they apply "
         "sustainability knowledge and skills (name, course and graduating year, a short summary, a link to a profile if available, "
         "and whether consent to be featured is likely). If your faculty tracks sustainability-related graduate outcomes in any way, "
         "please briefly describe how.",
         ""),
    12: ("Please provide up to two examples from 2026 of researchers in your faculty contributing to the University's own "
         "operational sustainability challenges (e.g. advising on campus energy, biodiversity, waste or buildings; applied projects; "
         "decision-support tools). For each, briefly describe the operational challenge, the researcher's contribution, the "
         "University team involved and the result.",
         "We will collect Campus Living Lab Accelerator Program (CLLAP) projects separately from CI&S – please include other examples only."),
    13: ("Does your faculty have any mechanism that supports ongoing collaboration between academic and professional staff on "
         "operational sustainability (e.g. a sustainability committee with mixed membership, seed funding, a coordination role or a "
         "regular forum)? If yes, briefly describe it and how often it operates. If not, please say so.",
         ""),
    14: ("For the collaborations described in TR1(a)–(b), please give a short assessment (a paragraph is sufficient) of their "
         "quality, relevance and continuity: are they ongoing, useful to the University's operational priorities, and producing outcomes?",
         ""),
    15: ("Has your faculty collected any feedback from participants (researchers or operational staff) on how well these "
         "collaboration mechanisms work and whether they are mutually beneficial? If yes, summarise key findings and any improvements "
         "identified; if not, please say so.",
         ""),
    16: ("Please provide one or two examples from 2026 of sustainability-related research from your faculty that has had impact "
         "beyond academia – on policy, industry, community, practice, cultural or institutional outcomes. For each, briefly describe "
         "the research, the partner or user, and the resulting change or influence, with links (e.g. Pursuit articles, policy "
         "submissions, reports).",
         "We will source University-wide research metrics (e.g. SDG-mapped publications via SciVal) from MRE – publication counts are not required."),
    19: ("Please provide one case study (up to 300 words) of a research project from your faculty that has contributed to a "
         "meaningful sustainability outcome: the research, stakeholders involved, pathway to impact and evidence of the outcome. "
         "This can be one of the examples from TR2(a), developed in more detail.",
         "We are also seeking examples from Melbourne Climate Futures, the Melbourne Biodiversity Institute and Strategic Communications."),
    20: ("Please describe how sustainability considerations are built into your faculty's research systems or support – e.g. "
         "research planning or approval processes, guidance, graduate researcher induction or training, grant development support. "
         "Please highlight any new or improved process introduced in 2026.",
         ""),
    21: ("Are there any efforts within the faculty to make research practices more sustainable (e.g. green lab certification, "
         "equipment sharing, reduced research travel, sustainable fieldwork, sustainable procurement of research consumables)? If so, "
         "please describe up to two examples and any evidence of uptake, resource savings or emissions reductions.",
         ""),
    22: ("Does your faculty have any evidence of change in sustainability-related research culture, capability or awareness – e.g. "
         "results from staff or graduate researcher surveys, or participation in capability-building events or training? Please "
         "provide a brief summary and the source. If no evidence is available, please say so.",
         ""),
    23: ("Drawing on your responses to TR3(a)–(c), please provide a short overall assessment (2–3 sentences) of your faculty's "
         "progress in embedding sustainable research practices: where progress is strongest, the main gaps, and priority actions for 2027.",
         ""),
    27: ("Please list up to three examples of 2026 communications from your faculty that showcase climate action, research or "
         "learning and aim to inform policy, sector or community decarbonisation (e.g. policy submissions, public events, media, "
         "reports). Include links, the intended audience and any reach data you have.",
         "This will also contribute to the University's mapping of decarbonisation-related research and education (due by 2028)."),
    29: ("If not already covered in CL2(a), please list any ongoing channels or programs (e.g. event series, newsletters, podcasts, "
         "policy engagement) through which your faculty regularly communicates climate research and learning to policy, sector or "
         "community audiences. You may respond 'see CL2(a)' if already covered.",
         "We will also engage Melbourne Climate Futures and the Melbourne Energy Institute."),
    52: ("Please provide one case study of an activity in your faculty (education, research, engagement or operations) where "
         "Indigenous knowledges informed a sustainability outcome. Describe the activity, how Indigenous knowledges were incorporated "
         "or co-created and with whom, and the result. Please confirm that appropriate permissions are in place for sharing.",
         ""),
    53: ("Please provide a case study of a living lab which operates within your faculty, including a description of its "
         "activities, who is involved (students, researchers, operational staff) and its outputs, outcomes or impact in 2026.",
         "Please note we will approach living labs funded through the Campus Living Lab Accelerator Program (CLLAP) separately."),
    54: ("Please list the key sustainability-focused communities of practice, internal networks or external partnerships active in "
         "your faculty in 2026 (name, partners, priority area), and provide a short description of one that has contributed to a "
         "Sustainability 2030 outcome.",
         ""),
    55: ("Please describe how sustainability is built into your faculty's governance and decision-making – e.g. a sustainability "
         "committee or Associate Dean (Sustainability) role, a faculty sustainability action plan, sustainability criteria in "
         "planning or approvals, and how progress is reported (e.g. to the Faculty Executive). Please link the action plan if one exists.",
         ""),
}


def faculty_request(row, sr25_ref, prev_status, prev_response, original, tailor=True, linked=False):
    """Compose the request text: opening line pointing to the end-2025 response, then the ask.
    tailor=False keeps the original S2030 wording (central units)."""
    req, note = TAILORED.get(row, (original, "")) if tailor else (original, "")
    who = "your faculty" if tailor else "your area"
    where = "click the ref in Column H to see it in full in Tab 3" if linked else "see Column J – click [+] above Columns I–K"
    parts = []
    if sr25_ref:
        if prev_response:
            st = f" and rated it '{prev_status}'" if prev_status else ""
            parts.append(f"In end-2025 reporting {who} reported on the related SP2030 target {sr25_ref}{st} "
                         f"({where}). Please build on this rather than repeating it, focusing on what is new in 2026.")
        else:
            parts.append(f"We did not receive an end-2025 response for the related SP2030 target {sr25_ref}, "
                         "so please include any relevant activity from across 2026.")
    parts.append(req)
    if note:
        parts.append(note)
    return "\n\n".join(parts)


# "How SST will use this data" – SR25 style, keyed by S2030 'Qual' worksheet row.
# 2026 is a transition year: no published traffic lights; responses also set the baseline for 2027 reporting.
USE = {
    8: "We will collate examples of innovative sustainability teaching and learning across faculties to inform the Exceptional education section of the 2026 Sustainability Report, and select one or two for short case studies.",
    9: "This is a new indicator. We will use responses to establish a baseline of how (and whether) faculties are understanding student sustainability self-efficacy, and report on early pilots in the 2026 Sustainability Report. A nil response is useful to us.",
    10: "We will combine faculty examples with central Wattle Fellowship and Melbourne Plus data to report on the range and reach of sustainability-focused experiential and co-curricular learning across the University.",
    11: "We will use nominated alumni to develop short graduate profiles for the 2026 Sustainability Report (subject to consent), and record any faculty approaches to tracking graduate outcomes to inform how this indicator is measured from 2027.",
    12: "We will report on examples of researchers helping solve the University's own operational sustainability challenges, alongside CLLAP projects reported by CI&S, and share them with operational teams to identify further collaboration.",
    13: "We will map existing academic–professional collaboration mechanisms across faculties to establish a baseline for this new indicator and identify models other faculties could adopt.",
    14: "We will use your assessment to describe the quality and continuity of research–operations collaboration in the 2026 report, and as baseline evidence for the interim target status assessment from 2027.",
    15: "We will use any feedback to understand whether collaboration mechanisms are working and to inform improvements. Where none has been collected, this helps us plan how to gather feedback from 2027.",
    16: "We will collate research impact examples across faculties to show the breadth of sustainability-related research impact in the Transformational research section of the 2026 Sustainability Report, complementing University-wide metrics from MRE.",
    18: "We will use institutional research output data (e.g. SDG-mapped publications via SciVal) to provide the University-wide quantitative picture of sustainability-related research in the 2026 Sustainability Report.",
    19: "We will select the strongest examples to develop as full research case studies for the 2026 Sustainability Report (we may contact you for more detail and images).",
    20: "We will report on how sustainability is being built into research systems and support across the University, and use responses to establish a baseline for this new target.",
    21: "We will report on ways the University is making research practices more sustainable in the 2026 Sustainability Report, and share practical examples (e.g. green labs) across faculties.",
    22: "We will use any survey or capability data to describe changes in sustainability-related research culture. Where no evidence exists, this helps us plan University-wide measurement (e.g. staff and graduate researcher surveys) from 2027.",
    23: "We will use faculty self-assessments to identify strengths, gaps and priority actions in embedding sustainable research practices, to inform internal management reporting and the Transformational research working group.",
    26: "For internal planning: we will record the status of the planned review of Climate Leadership targets (due by 2030).",
    27: "We will use these examples to show how the University communicates climate research and learning to external audiences, and as an input to the institutional decarbonisation research and education mapping due by 2028.",
    29: "We will report on the University's ongoing climate communications in the Climate leadership section of the 2026 Sustainability Report, alongside contributions from Melbourne Climate Futures and the Melbourne Energy Institute.",
    30: "We will report the University's climate-change resilience maturity rating, key gaps and priority actions in the 2026 Sustainability Report (building on the Climate Change Preparedness Framework assessment reported in 2025).",
    32: "For internal management reporting only: we will track the size, health and diversity of remnant vegetation at Dookie and Creswick. This information will not be published.",
    33: "We will report progress on developing the biodiversity quality measurement framework, which will underpin how the biodiversity target is assessed from 2027.",
    34: "We will report the definition and baseline for campus ecological connectivity in the 2026 Sustainability Report, to support future tracking against the biodiversity target.",
    35: "We will report on the baseline nature footprint of the University's supply chain in the Nature and biodiversity section of the 2026 Sustainability Report.",
    36: "We will describe the framework for monitoring supply-chain nature impacts, and use it to plan how this target is reported from 2027.",
    37: "We will report on the high-risk procurement categories or suppliers prioritised for nature-impact management, and related actions, in the 2026 Sustainability Report.",
    38: "We will provide a high-level overview of AI's estimated operational impacts and opportunities in 2026, ahead of more detailed measurement from 2027.",
    39: "We will select examples of sustainability and AI-related education, research and collaboration as case studies for the new Responsible AI section of the 2026 Sustainability Report.",
    40: "We will use case study evidence to show how sustainability maturity is increasing in estate and infrastructure planning, development and operations in the 2026 Sustainability Report.",
    41: "We will use estimated sustainability impacts of the capital plan to inform progress reporting towards the emissions reduction target and other priority areas (to the extent information is available).",
    42: "We will report on sustainability activities and impacts of major estate and infrastructure projects and programs in the Estate and infrastructure section of the 2026 Sustainability Report.",
    43: "We will report on how Indigenous approaches to sustainability are informing estate planning and development, with examples where permission to share is confirmed.",
    46: "We will use case studies (and any maturity score or training data) to show improvement in responsible procurement processes and systems in the 2026 Sustainability Report.",
    49: "We will draw on the annual Modern Slavery Statement and case studies to report on the University's approach to modern slavery and human rights risks (timing to be aligned with the Statement).",
    51: "We will report on sustainability-related investment portfolio metrics and the University's compliance with related obligations (e.g. UNPRI) in the Responsible investments section of the 2026 Sustainability Report.",
    52: "We will report on the ways Indigenous knowledges inform sustainability activities across the University, and develop selected examples into case studies with appropriate permissions.",
    53: "We will share examples of living labs from across the University in the 2026 Sustainability Report, alongside CLLAP projects reported by CI&S.",
    54: "We will report on internal engagement and partnerships contributing to Sustainability 2030 outcomes, and update our record of sustainability-focused communities of practice and contacts across the University.",
    55: "We will report on how sustainability is integrated into governance and decision-making across faculties and portfolios, and use responses as a baseline for the governance enabler from 2027.",
}
