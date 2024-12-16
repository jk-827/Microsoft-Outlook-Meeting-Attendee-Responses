"""
Run this script and paste in the "Copy Status to Clipboard" contents from a Microsoft Outlook Meeting Invite.
Returns nicely formatted details of names with their response types. Output is ordered alphabetically by first name and
grouped by response type.
"""


def outlook_meeting_formatter():
    # Get the required input.
    lines = []
    print("Copy and paste in the 'Copy Status to Clipboard' content. Then press Enter:\n")
    while True:
        line = input()
        if not line:
            break
        lines.append(line)

    # Break up each line and create lists of names (so we can find the longest name) and people (list of lists).
    # Exclude certain names if required such as the Meeting Organizer.
    people = []
    names = []
    for line in lines:
        person = line.split("\t")
        excluded_names = ["Name"]
        if person[0] not in excluded_names:
            people.append(person)  # people is a list of all the data associated with each name
            names.append(person[0])  # names is a list of just the names

    # Check each person and print their response status
    response_lists = {
        'Accepted': [],
        'Declined': [],
        'None': [],
        'Tentative': []
    }

    # Use the longest name to calculate how many spaces to add so print statements format nicely
    max_length = len(max(names, key=len)) + 1

    # Create a list for each response type
    for person in people:
        name, status = person[0], person[2]
        spaces = max_length - len(name)
        if status in response_lists:
            response_lists[status].append([name, status, spaces])
        else:
            continue

    # Create the final list ready for output. Accepted people will be numbered.
    final_list = []
    place_num = 1

    # Add numbers and adjust spaces for Accepted people
    for name, status, spaces in sorted(response_lists['Accepted']):
        if len(str(place_num)) > 1:
            # More than 1 digit, so normal spaces
            x = ("#" + str(place_num) + " - " + name + " "*spaces + "--> ACCEPTED\n")
        else:
            # If place num is only 1 digit long, add one extra space
            x = ("#" + str(place_num) + "  - " + name + " " * spaces + "--> ACCEPTED\n")
        final_list.append(x)
        place_num += 1

    # Add a dashed line as a separator
    full_dashed_line = ("-" * len(x) + "\n")
    final_list.append(full_dashed_line)

    # Next add Tentative people
    for name, status, spaces in sorted(response_lists['Tentative']):
        x = (name + " " * spaces + "          --> TENTATIVE\n")
        final_list.append(x)

    # Next add Declined people
    for name, status, spaces in sorted(response_lists['Declined']):
        x = ('X - ' + name + " " * spaces + "  --> DECLINED\n")
        final_list.append(x)

    # Next add No Response people
    for name, status, spaces in sorted(response_lists['None']):
        x = ('? - ' + name + " " * spaces + "     --> NO RESPONSE\n")
        final_list.append(x)

    # Add a dashed line as a separator
    final_list.append(full_dashed_line)

    # Add the final counts for each response type
    counters_dict = {'ACCEPTED': len(response_lists['Accepted']),
                     'TENTATIVELY ACCEPTED': len(response_lists['Tentative']),
                     'DECLINED': len(response_lists['Declined']),
                     'NOT RESPONDED': len(response_lists['None'])}

    for key in counters_dict:
        if counters_dict[key] == 1:
            x = (str(counters_dict[key]) + ' person has ' + key + "\n")
        else:
            x = (str(counters_dict[key]) + ' people have ' + key + "\n")
        final_list.append(x)

    # Add a dashed line as a separator
    final_list.append(full_dashed_line)

    # Write to a file
    with open("Meeting Attendees with Status.txt", "w") as f:
        f.writelines(final_list)
    f.close()

    return {'Final List': final_list, 'Accepted List': response_lists['Accepted']}


def get_accepted_names(accepted_list):
    results_list = ''
    for item in accepted_list:
        results_list += f"{item[0]};"
    return f"{results_list}"


if __name__ == '__main__':
    res = outlook_meeting_formatter()

    with open("Meeting Attendees with Status.txt", "r") as f:
        lines = (f.readlines())
        for line in lines:
            print(line, end='')

    print('People who are attending to email:')
    print(get_accepted_names(res['Accepted List']))
