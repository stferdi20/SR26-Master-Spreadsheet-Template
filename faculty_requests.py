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


def faculty_request(row, sr25_ref, prev_status, prev_response, original):
    """Compose the faculty-specific request text."""
    req, note = TAILORED.get(row, (original, ""))
    parts = []
    if sr25_ref:
        if prev_response:
            st = f" and rated it '{prev_status}'" if prev_status else ""
            parts.append(f"In end-2025 reporting your faculty reported on the related SP2030 target {sr25_ref}{st} "
                         "(see Column J – click [+] above Columns H–K). Please build on this rather than repeating it, "
                         "focusing on what is new in 2026.")
        else:
            parts.append(f"We did not receive an end-2025 response for the related SP2030 target {sr25_ref}, "
                         "so please include any relevant activity from across 2026.")
    parts.append(req)
    if note:
        parts.append(note)
    return "\n\n".join(parts)
