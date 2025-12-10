# Import Libraries 
import streamlit as st
import requests
import random
import time
from email_validator import validate_email, EmailNotValidError
from captcha.image import ImageCaptcha
from streamlit_js_eval import streamlit_js_eval

# -----------------------------
# Page setup
# -----------------------------
st.set_page_config(
    page_title="Contact",
    page_icon="✉️",
    layout="wide",
)

st.sidebar.header("Contact")

with st.sidebar:
    st.markdown("### Navigation")
    st.caption("Use the menu above to switch between pages.")
    st.caption("You are on the **Contact** page. Reach out if you need help or have questions about the app.")

    with st.expander("ℹ️ How this page works", expanded=False):
        st.markdown(
            """
Use this page to get in touch regarding:

- App support or troubleshooting  
- Feature requests  
- Bug reports  
- Questions about the cleaning or comparison logic  

I do my best to respond to inquiries as quickly as possible, but please note that replies may not be immediate.
            """
        )

# -----------------------------
# Formspree endpoint and CAPTCHA setup
# -----------------------------
try:
    FORMSPREE_ENDPOINT = st.secrets["formspree_endpoint"]
except KeyError:
    st.error("Form endpoint is not configured. Please set 'formspree_endpoint' in Streamlit secrets.")
    st.stop()

# Define CAPTCHA character set
CAPTCHA_OPTIONS = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"

# Generate CAPTCHA
def generate_captcha():
    captcha_text = "".join(random.choices(CAPTCHA_OPTIONS, k=6))
    image = ImageCaptcha(width=350, height=100).generate_image(captcha_text)
    return captcha_text, image

# Initialize CAPTCHA in session
if "captcha_text" not in st.session_state:
    st.session_state.captcha_text, st.session_state.captcha_image = generate_captcha()

# -----------------------------
# Layout
# -----------------------------
col1, col2, col3, col4 = st.columns([3, 0.25, 1, 0.25])

# CAPTCHA section
with col3:
    st.markdown("##### CAPTCHA verification")
    st.markdown(
        '<p style="text-align: justify; font-size: 12px;">CAPTCHAs help protect this form from automated spam. '
        'Please type the characters you see in the image below.</p>',
        unsafe_allow_html=True,
    )

    captcha_placeholder = st.empty()
    captcha_placeholder.image(st.session_state.captcha_image, use_container_width=True)

    if st.button("🔁 Refresh CAPTCHA"):
        st.session_state.captcha_text, st.session_state.captcha_image = generate_captcha()
        captcha_placeholder.image(st.session_state.captcha_image, use_container_width=True)

    captcha_input = st.text_input("Enter CAPTCHA")

    st.write("")
    st.write("")

    st.markdown("#### Interested in **:rainbow[more?]**")
    st.markdown(
        "You can see some of my other work on my portfolio "
        "[here](https://erin-weiss.github.io)."
    )

# Contact form section
with col1:
    st.header("📬 Contact me")
    st.write(
        "If you have questions about this app, run into issues, or have ideas for improvements, "
        "you can use this form to send a message."
    )

    name = st.text_input("Your name*", max_chars=100)
    email = st.text_input("Your email*", max_chars=100)
    subject = st.text_input("Subject*", max_chars=150)
    message = st.text_area("Your message*", height=200)

    st.markdown('<p style="font-size: 13px;">* Required fields</p>', unsafe_allow_html=True)

    if st.button("Send message", type="primary"):
        if not name or not email or not subject or not message:
            st.error("Please fill out all required fields before sending your message.")
        else:
            try:
                # Validate email
                validate_email(email, check_deliverability=True)

                # CAPTCHA check
                if captcha_input.upper() == st.session_state.captcha_text:
                    data = {
                        "name": name,
                        "email": email,
                        "subject": subject,
                        "message": message
                    }

                    response = requests.post(
                        FORMSPREE_ENDPOINT,
                        data=data,
                        headers={"Accept": "application/json"}
                    )

                    if response.status_code == 200:
                        st.success("✅ Your message has been sent successfully.")

                        # Reset CAPTCHA
                        st.session_state.captcha_text, st.session_state.captcha_image = generate_captcha()
                        captcha_placeholder.image(st.session_state.captcha_image, use_container_width=True)

                        # Brief pause and reload to clear the form
                        time.sleep(2)
                        streamlit_js_eval(js_expressions="parent.window.location.reload()")
                    else:
                        st.error("⚠️ The message could not be sent. Please try again.")
                        st.text(f"Error: {response.status_code} - {response.text}")
                else:
                    st.error("❌ CAPTCHA did not match. Please try again.")
            except EmailNotValidError as e:
                st.error(f"Invalid email address. {e}")
