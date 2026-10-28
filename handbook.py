"""Company Employee Handbook: types and degrees of offenses (from Sanction.pdf).

Each entry is (section number, text, degree). The section numbers and wording follow the handbook, including its
own gaps and one repeated number (3.2.5 appears twice). Do not "fix" them: the Word file cites these as written.
"""

import re

CATEGORIES = {
    "1": "Attendance and Punctuality",
    "2": "Behavior & Conduct",
    "3": "Good Morals and Work Ethics",
    "4": "Negligence in the Performance of Duty",
    "5": "Company Interest/ Insubordination",
    "6": "Company Funds and Property",
    "7": "Sanitation, Safety & Security at Work",
    "8": "Securing or Divulging Confidential Information",
}

SECTIONS = [
    # 1. ATTENDANCE AND PUNCTUALITY
    ("1.1.1", "Tardiness: one late in a month", "1st"),
    ("1.1.2", "Tardiness: two times late in a month", "3rd"),
    ("1.1.3", "Tardiness: three times late in a month", "4th"),
    ("1.1.4", "Tardiness: four times late in a month", "5th"),
    ("1.1.5", "Tardiness: five times late in a month", "5th"),
    ("1.1.6", "Tardiness: six times late in a month", "7th"),
    ("1.1.7", "Tardiness: seven times or more", "7th"),
    ("1.2.1", "Absent without official leave (AWOL) or absent without notice or approval: one (1) day", "1st"),
    ("1.2.2", "Absent without official leave (AWOL) or absent without notice or approval: two (2) days", "3rd"),
    ("1.2.3", "Absent without official leave (AWOL) or absent without notice or approval: three (3) days", "4th"),
    ("1.2.4", "Absent without official leave (AWOL) or absent without notice or approval: four (4) days", "5th"),
    ("1.2.5", "Absent without official leave (AWOL) or absent without notice or approval: five (5) days", "5th"),
    ("1.2.6", "Absent without official leave (AWOL) or absent without notice or approval: six (6) days", "7th"),
    ("1.2.7", "Absent without official leave (AWOL) or absent without notice or approval: seven (7) days", "7th"),
    ("1.3.1", "Unauthorized under-time within a month: one (1) day", "1st"),
    ("1.3.2", "Unauthorized under-time within a month: two (2) days", "3rd"),
    ("1.3.3", "Unauthorized under-time within a month: three (3) days", "4th"),
    ("1.3.4", "Unauthorized under-time within a month: four (4) days", "5th"),
    ("1.3.5", "Unauthorized under-time within a month: five (5) days", "5th"),
    ("1.3.6", "Unauthorized under-time within a month: six (6) days", "7th"),
    ("1.3.7", "Unauthorized under-time within a month: seven (7) days", "7th"),
    ("1.6", "Misuse of Policy on break period (over breaking)", "1st"),
    # 2. BEHAVIOR & CONDUCT
    ("2.1", "Insubordination or wilful refusal to obey company rules and Regulation or lawful orders of immediate superior", "4th-7th"),
    ("2.2", "Endangering the safety of other employee through carelessness or negligence", "2nd-5th"),
    ("2.3", "Causing physical injuries to co-worker on account of carelessness or negligence", "4th-5th"),
    ("2.4", "Rumormongering, unnecessary disclosure of somebody else's affairs to other or deliberate distortion of facts or statement in such a way as to enhance one's status or reputation or discredit, embarrasses another employee", "5th-7th"),
    ("2.5", "Provoking, threatening, intimidating or instigating a fight and fighting within the company premises and outside the Company", "5th-7th"),
    ("2.6", "Inflicting physical injuries, coercing or harassing a co-employee, supervisor or anybody within the company premises", "5th-7th"),
    ("2.7", "Threatening the life of a company official, co-employee or anybody on the company or within company premises", "5th-7th"),
    ("2.8", "Intriguing against the honour of company officials and clients", "5th-7th"),
    # 3. GOOD MORALS AND WORK ETHICS
    ("3.1.1", "Using indecent, abusive, derogatory and/or insidious words verbally or in writing or in addressing another person or any of its official within the company premises", "5th"),
    ("3.1.2", "Using profane or obscene language in addressing another person within company premises", "5th"),
    ("3.1.3", "Uttering words, performing acts or making gestures which are manifestly insulting, disrespectful and grossly discourteous to clients, guests and visitors", "5th"),
    ("3.1.4", "Blatant use of disrespectful language and manners against company officials within company premises", "5th"),
    ("3.2.1", "Committing fraud and deception in dealing with co-employees and customers", "5th"),
    ("3.2.2", "Misappropriation of company funds for personal use", "7th"),
    ("3.2.3", "Promoting, engaging in and practicing usury", "7th"),
    ("3.2.4", "Offering or accepting bribes or inducing others to violate Company rules", "7th"),
    ("3.2.5", "Gambling, placing or collecting bets or participating in any game of chance during official working hours or within company premises", "7th"),
    ("3.2.5", "Stealing from the company, its officers, fellow employees, or superiors", "7th"),
    ("3.3.1", "Giving false information in the application form of personnel records such as birthdates, civil status, employment dates, experiences, educational attainment, etc", "7th"),
    ("3.3.2", "Tampering, altering or falsifying of official papers and/or documents whether public or private, and falsifying signatures of company employees, officials and executives", "7th"),
    ("3.3.3", "Giving false testimony during official company investigation", "7th"),
    ("3.4.1", "Possession and/or use of any mood altering substances/illegal or prohibited drugs within company premises", "7th"),
    ("3.4.2", "Entering company premises while inebriated and/or drinking liquor or other intoxicating beverages within the company premises", "7th"),
    ("3.5.1", "Engaging in immoral or unethical acts or practices or offering immoral or unethical goods or services to employees or officers", "7th"),
    ("3.5.2", "Sexual harassment, engaging in physical or verbal abuse of co-employees of a sexual nature", "7th"),
    ("3.5.3", "Engaging in a scandalous relationship with another employee while validly married; illicit relationship with co-employees", "7th"),
    ("3.6", "Abuse of status, power of discretion by superiors; use of force, violence or manhandling which would injure, degrade or humiliate a subordinate", "7th"),
    ("3.7", "Conviction of any crime or felony involving moral turpitude", "7th"),
    ("3.8", "Omission of acts which are grossly against public morals not embraced by these provisions", "7th"),
    # 4. NEGLIGENCE IN THE PERFORMANCE OF DUTY
    ("4.1", "Gross negligence in the performance of duties and responsibilities (penalty depends on the damage caused to the company)", "4th-5th"),
    ("4.2", "Leaving work assignment before quitting time without permission (penalty depends on the damage caused to the company)", "5th-7th"),
    ("4.3", "Wasting time or loitering on company time, stretching break period, and idle chatting using company time to make personal calls", "4th"),
    ("4.4", "Engaging in horseplay or other playful acts which are time consuming and may be harmful to person or property", "5th"),
    ("4.5", "Engaging in an unauthorized/unofficial undertaking during working hours", "5th"),
    ("4.6", "Sleeping while on duty (penalty depends on the damage caused to the company)", "1st-7th"),
    ("4.7", "Failure to cooperate with other employees, unjustified refusal to extend assistance to fellow employee when necessary in the operations", "1st"),
    ("4.8", "Failure to perform duties and responsibilities including but not limited to improper or non-accomplishment of reports, documents and non-compliance with the standard operational procedure", "5th"),
    ("4.9", "Refusing or failing to report for overtime work after having been scheduled and properly informed without justifiable reasons", "5th"),
    ("4.10", "Wilful condonement, toleration or participation of the Supervisor/s in illegal activities of their subordinates", "7th"),
    ("4.11", "Other acts committed by a subordinate for reasons directly connected with the superior's discharge of official duties which is clearly prejudicial to the company's interest", "7th"),
    # 5. COMPANY INTEREST / INSUBORDINATION
    ("5.1", "Wilful failure to obey lawful orders of immediate superior or department head", "5th-7th"),
    ("5.2", "Wilful disregard to obey established company policies, rules and procedure", "5th-7th"),
    ("5.3", "Wilful failure to perform assigned tasks", "5th-7th"),
    ("5.4", "Wilful refusal to accept or acknowledge written notices or official communications from the company without just cause", "5th-7th"),
    ("5.5", "Offering or accepting anything of value for a job or for a favourable condition of employment in exchange for the company's reputation and/or integrity and secrecy", "5th-7th"),
    ("5.7", "Unjustified refusal to submit to an official investigation authorized by the company or giving false testimony during such official investigation", "5th-7th"),
    ("5.8", "Unauthorized use of company name to obtain personal benefits", "5th-7th"),
    ("5.9", "Distributing written or printed matters detrimental to the interest of the company", "5th-7th"),
    ("5.10", "Deliberately slowing down, hindering or limiting production or intimidating, coercing or inducing other employees to do so", "5th-7th"),
    ("5.11", "Inciting or participating in concerted work stoppage, slow down, mass leave, sit-down strike or other similar disruptive activities", "5th-7th"),
    ("5.12", "Making false or malicious statements about the company and/or its products/services or any of its employees", "5th-7th"),
    ("5.13", "Damaging or jeopardizing the company's interest through acts or omissions which affect company clients/official/visitors or other business associates of the management", "5th-7th"),
    ("5.14", "Engaging in any activity which is directly competitive with the company's business or any part of its operation while still connected in the company", "5th-7th"),
    ("5.15", "Forging, falsifying or altering official documents in such a way to mislead the users thereof", "5th-7th"),
    ("5.16", "Breach by employees of the trust and confidence reposed by Management or a company representative", "5th-7th"),
    ("5.17", "Committing any other act of dishonesty or other culpable acts or omissions not embraced by other provisions which cause damage to company interest", "5th-7th"),
    # 6. COMPANY FUNDS AND PROPERTY
    ("6.1.1", "Vandalism, posting or removing anything from the bulletin boards or other places for notices without authorization", "4th"),
    ("6.1.2", "Unauthorized, reckless, negligent use of machine & equipment & vehicles", "5th"),
    ("6.1.3", "Operating, using, meddling with or impeding the proper use of tools, equipment or machine which has not been assigned or trained", "5th"),
    ("6.1.5", "Wilful disregard of, or violating established company safety rules", "4th"),
    ("6.1.6", "Removing safety devices from machinery, equipment & vehicle", "5th"),
    ("6.1.7", "Removing company property from premises, concealing or deliberately misplacing company property without justifiable purpose", "5th"),
    ("6.1.8", "Tampering of switches, controls, installations of machines, equipments and vehicles", "5th-7th"),
    ("6.1.9", "Using company properties, materials or equipment for personal use (e.g. telephone)", "5th"),
    ("6.1.10", "Abusive use of company properties, materials and supplies", "5th"),
    ("6.1.11", "Taking home company property supplies and materials without duly accomplished form issued by the HRD/Admin and IT Technician", "5th"),
    ("6.1.12", "Driving the company vehicle without license, with expired license, or with fake license", "5th-7th"),
    ("6.1.13", "Culpable carelessness, negligence, failure to follow specific instruction or established procedures resulting to loss or damage to company property", "7th"),
    ("6.1.14", "Unauthorized possession and/or master key which can open doors, desk, drawers and cabinets as well as vehicles", "7th"),
    ("6.1.15", "Damaging company property through misuse or negligence", "7th"),
    ("6.2.1", "Substituting company materials or equipment without authorization", "7th"),
    ("6.2.2", "Stealing, misappropriating or embezzling of company funds (this includes qualified theft)", "7th"),
    ("6.2.3", "Obtaining or attempting to obtain materials fraudulently or through fraudulent orders and/or requisition", "7th"),
    ("6.2.4", "Fraudulent or illegitimate withdrawal, acquisition or release to unauthorized person of company funds or property", "7th"),
    ("6.3", "Committing an act of sabotage (deliberate destruction of company property or hindrance/obstruction of normal operations)", "7th"),
    ("6.4", "Committing other acts of dishonesty, deceit or other irregularities not embraced by other provision which cause loss or damage to company property", "7th"),
    # 7. SANITATION, SAFETY & SECURITY AT WORK
    ("7.1.1", "Poor housekeeping of assigned area and other facilities within the company premises", "1st"),
    ("7.1.2", "Neglect of personal cleanliness hygiene and grooming", "4th"),
    ("7.1.3", "Creating or contributing to unsanitary conditions like leaving spoiled dishes in the pantry, spitting, storing food inside drawers and littering", "1st"),
    ("7.2.1", "Failure to wear identification card during company time/working hours", "1st"),
    ("7.2.2", "Failure to comply with physical examination requirements of the HR", ""),
    ("7.2.3", "Failure to meet or observe security requirement or regulation of the company", "1st"),
    ("7.2.4", "Refusal to submit oneself for security inspection", "4th"),
    ("7.2.5", "Participating in loud and heated arguments during official working hours and/or within company premises", "5th"),
    ("7.2.6", "Malingering, not reporting for work under the pretext of sickness", "5th"),
    ("7.2.7", "Unauthorized overnight stay within company premises", "5th"),
    ("7.2.8", "Smoking in prohibited areas", "5th"),
    ("7.2.9", "Entering unauthorized/restricted areas without permission", "5th"),
    ("7.2.10", "Carrying firearms or any deadly weapon or explosives, defined by laws, without permission from superior or authorities at any time within company premises with or without intent to do bodily harm to any person", "7th"),
    ("7.2.11", "Failure to report affliction of a contagious disease to company authorities", "5th"),
    ("7.3", "Omission of other acts grossly affecting the sanitation, safety and security at work", "7th"),
    # 8. SECURING OR DIVULGING CONFIDENTIAL INFORMATION
    ("8.1", "Transmitting confidential information concerning company operations, formulas, products and the like without permission", "7th"),
    ("8.2", "Disclosing, giving away or using restricted or classified company matters/information without proper authorization", "7th"),
    ("8.3", "Transmitting confidential information concerning company operations, formulas, products and the like without permission", "7th"),
    ("8.2.1", "If the employee is responsible for or has access to data", "7th"),
]


def category_of(section_no):
    return section_no.split(".")[0]


def label(i):
    """Text shown in the picker, e.g. '4.1 - Gross negligence ... (4th-5th)'. i is an index into SECTIONS."""
    no, text, deg = SECTIONS[i]
    short = text if len(text) <= 90 else text[:87].rstrip() + "..."
    return f"{no}  {short}" + (f"  ({deg})" if deg else "")


def citation(i):
    """The line written in the Word file: 'Section 4.1. Gross negligence ...; (4th-5th Degree)'."""
    no, text, deg = SECTIONS[i]
    body = re.sub(r"\s*\(penalty depends[^)]*\)", "", text, flags=re.I).rstrip(" .;")
    return f"Section {no}. {body}" + (f"; ({deg.replace('-', chr(0x2013))} Degree)" if deg else "")


def heading(indices):
    """Category names of the chosen sections, in order, joined with 'and' (as in the template)."""
    names = []
    for i in indices:
        n = CATEGORIES[category_of(SECTIONS[i][0])]
        if n not in names:
            names.append(n)
    if not names:
        return ""
    # "Company Interest/ Insubordination" -> keep the handbook's own wording
    return " and ".join(names) + "."


def compact_list():
    """One short line per section for the AI: 'index|number|short text|degree'."""
    out = []
    for i, (no, text, deg) in enumerate(SECTIONS):
        out.append(f"{i}|{no}|{text[:110]}|{deg}")
    return "\n".join(out)
