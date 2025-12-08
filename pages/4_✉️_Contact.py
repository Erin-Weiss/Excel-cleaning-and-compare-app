# Import Libraries 
import streamlit as st
import requests

import random
import os
import time
from email_validator import validate_email, EmailNotValidError
from captcha.image import ImageCaptcha
from PIL import Image
from streamlit_js_eval import streamlit_js_eval

# Set up page
st.set_page_config(
    page_title="Contact",
    page_icon="✉️",
)

st.sidebar.header("Contact")

def wide_space_default():
    st.set_page_config(layout="wide")

wide_space_default()

with st.sidebar:
    st.markdown("### Navigation")
    st.caption("Use the menu above to switch between pages.")
    st.caption("You're on the **Contact** page. Reach out if you need help or have questions about the app.")

    with st.expander("ℹ️ How this page works", expanded=False):
        st.markdown(
            """
            Use this page to get in touch regarding:

            - App support or troubleshooting  
            - Feature requests  
            - Bug reports  
            - Questions about cleaning or comparison logic

            I do my best to respond to inquiries as quickly as possible, but please note that replies may not be immediate.
            """
        )



# Formspree endpoint
FORMSPREE_ENDPOINT = "https://formspree.io/f/mkgzvwle"

# Set wide layout
st.set_page_config(layout="wide")

# Define CAPTCHA character set
CAPTCHA_OPTIONS = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"

# Generate CAPTCHA
def generate_captcha():
    captcha_text = "".join(random.choices(CAPTCHA_OPTIONS, k=6))
    image = ImageCaptcha(width=350, height=100).generate_image(captcha_text)
    return captcha_text, image

# Initialize CAPTCHA in session
if 'captcha_text' not in st.session_state:
    st.session_state.captcha_text, st.session_state.captcha_image = generate_captcha()

# Layout
col1, col2, col3, col4 = st.columns([3, 0.25, 1, 0.25])

# CAPTCHA section
with col3:
    st.markdown("##### CAPTCHA Verification")
    st.markdown(
        '<p style="text-align: justify; font-size: 12px;">CAPTCHAs help protect this form from spam. Please enter the characters below.</p>',
        unsafe_allow_html=True,
    )

    captcha_placeholder = st.empty()
    captcha_placeholder.image(st.session_state.captcha_image, use_container_width=True)

    if st.button("🔁 Refresh CAPTCHA"):
        st.session_state.captcha_text, st.session_state.captcha_image = generate_captcha()
        captcha_placeholder.image(st.session_state.captcha_image, use_container_width=True)

    captcha_input = st.text_input("Enter CAPTCHA")

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("#### Interested in **:rainbow[More?]**")
    st.markdown("You can see some of my other work on my portfolio [here](https://erin-weiss.github.io)")

# Contact form section
with col1:
    st.header("📬 Contact Me")
    name = st.text_input("Your name*", max_chars=100)
    email = st.text_input("Your email*", max_chars=100)
    subject = st.text_input("Subject*", max_chars=150)
    message = st.text_area("Your message*", height=200)

    st.markdown('<p style="font-size: 13px;">* Required fields</p>', unsafe_allow_html=True)

    if st.button("Send Message", type="primary"):
        if not name or not email or not subject or not message:
            st.error("Please fill out all required fields.")
        else:
            try:
                # Validate email
                validate_email(email, check_deliverability=True)

                # CAPTCHA check
                if captcha_input.upper() == st.session_state.captcha_text:
                    # Submit form data to Formspree
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
                        st.success("✅ Message sent successfully!")

                        # Reset CAPTCHA
                        st.session_state.captcha_text, st.session_state.captcha_image = generate_captcha()
                        captcha_placeholder.image(st.session_state.captcha_image, use_container_width=True)

                        time.sleep(2)
                        streamlit_js_eval(js_expressions="parent.window.location.reload()")
                    else:
                        st.error("⚠️ Failed to send. Please try again.")
                        st.text(f"Error: {response.status_code} - {response.text}")
                else:
                    st.error("❌ CAPTCHA did not match.")
            except EmailNotValidError as e:
                st.error(f"Invalid email address. {e}")
