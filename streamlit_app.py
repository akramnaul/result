
workbook = get_workbook()

with st.sidebar:
    st.subheader("Faculty / Staff Access")
    if st.session_state.authenticated_email:
        st.success("Email verified against FacultyStaffList.")
        st.caption(st.session_state.authenticated_email)
        if st.button("Sign out / Change email", use_container_width=True):
            st.session_state.authenticated_email = ""
            st.session_state.email_checked = False
            st.session_state.roll_number = ""
            st.rerun()
    else:
        with st.form("email_authentication_form"):
            email_input = st.text_input("Faculty/staff email address", placeholder="name@example.com")
            email_submit = st.form_submit_button("Verify Email", use_container_width=True)

        if email_submit:
            if not email_input.strip():
                st.error("Please enter your email address.")
            elif email_is_authorized(workbook, email_input):
                st.session_state.authenticated_email = email_input.strip()
                st.session_state.email_checked = True
                st.success("Email address verified.")
                st.rerun()
            else:
                st.error("Email address not found. Please try again with the correct faculty/staff email.")

if not st.session_state.authenticated_email:
    st.info("Please verify your faculty/staff email address in the left panel to continue.")
    st.stop()

st.subheader("Find Student Results")
with st.form("roll_search_form"):
    roll_input = st.text_input(
        "Student Roll Number",
        value=st.session_state.roll_number,
        placeholder="Enter the student's roll number",
    )
    roll_submit = st.form_submit_button("Show All Available Monthly Results", type="primary")

if roll_submit:
    if not normalize_roll(roll_input):
        st.error("Please enter a student roll number.")
    else:
        st.session_state.roll_number = roll_input.strip()

if not st.session_state.roll_number:
    st.stop()

requested_roll = st.session_state.roll_number
result_sheets = sorted(
    [
        name for name in workbook.sheetnames
        if name.casefold() != FACULTY_SHEET.casefold() and is_monthly_result_sheet(name)
    ],
    key=month_sort_key,
)

if not result_sheets:
    st.warning("No monthly result worksheets were found. Expected names such as Result-26-09.")
    st.stop()

matches = []
for sheet_name in result_sheets:
    ws = workbook[sheet_name]
    result = get_month_result(ws, requested_roll)
    if result is not None:
        matches.append((sheet_name, result))

if not matches:
    st.error(
        f"Roll number '{requested_roll}' was not found in any available monthly result sheet. "
        "Please check the roll number and try again."
    )
    st.stop()

st.success(f"Found {len(matches)} monthly result(s) for roll number {requested_roll}.")
st.caption(f"Signed in as: {st.session_state.authenticated_email}")

for sheet_name, result in matches:
    with st.container(border=True):
        st.subheader(f"📅 {display_month(sheet_name)}")
        st.caption(f"Worksheet: {sheet_name}")

        left, right = st.columns([1, 1], gap="large")

        with left:
            st.markdown("**Student Information**")
            st.text_input("Roll Number", value=result["roll"], disabled=True, key=f"{sheet_name}_roll")
            st.text_input("Student Name", value=result["name"], disabled=True, key=f"{sheet_name}_name")
            st.text_input(
                "Attendance Percentage",
                value=result["attendance"],
                disabled=True,
                key=f"{sheet_name}_attendance",
            )

            st.markdown("**Subjects**")
            for i, (subject_name, _marks) in enumerate(result["subjects"], start=1):
                shown_name = subject_name if subject_name else f"Subject {i:02d}"
                st.text_input(
                    f"Subject {i:02d}",
                    value=shown_name,
                    disabled=True,
                    key=f"{sheet_name}_subject_name_{i}",
                )

        with right:
            st.markdown("**Attendance and Marks**")
            st.text_input(
                "Attendance (%)",
                value=result["attendance"],
                disabled=True,
                key=f"{sheet_name}_attendance_right",
            )
            st.text_input("Student Name", value=result["name"], disabled=True, key=f"{sheet_name}_name_right")

            st.markdown("**Subject Marks**")
            for i, (_subject_name, marks) in enumerate(result["subjects"], start=1):
                st.text_input(
                    f"Marks — Subject {i:02d}",
                    value=marks,
                    disabled=True,
                    key=f"{sheet_name}_marks_{i}",
                )
