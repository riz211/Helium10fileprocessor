"""Tutorial module for guiding new users through the application."""
import streamlit as st
from typing import Dict, Any, Optional

class TutorialGuide:
    """Manages the tutorial state and content for the application."""
    
    def __init__(self):
        """Initialize the tutorial guide with steps and state management."""
        if 'tutorial_active' not in st.session_state:
            st.session_state.tutorial_active = False
        if 'tutorial_step' not in st.session_state:
            st.session_state.tutorial_step = 0
            
    def get_tutorial_content(self) -> Dict[int, Dict[str, Any]]:
        """Return the content for each tutorial step."""
        return {
            0: {
                'title': 'Welcome to Helium 10 File Processor! 👋',
                'content': """
                This tutorial will guide you through the main features of the application.
                Click 'Next' to begin your journey!
                """,
                'highlight': None
            },
            1: {
                'title': 'Upload Files 📁',
                'content': """
                Start by uploading your Helium 10 Excel files using the file uploader.
                The app accepts multiple files at once!
                
                Key points:
                - Accepts .xlsx files
                - Can process multiple files simultaneously
                - Automatically combines data from all sheets
                """,
                'highlight': 'file_uploader'
            },
            2: {
                'title': 'Manage Blocked Items ⛔',
                'content': """
                Use the sidebar to manage blocked brands and product IDs:
                
                1. Switch between Brands and Product IDs
                2. Add items individually or bulk upload
                3. View and export your blocked items list
                """,
                'highlight': 'sidebar'
            },
            3: {
                'title': 'Process and Review Data 📊',
                'content': """
                After uploading files, the app will:
                
                1. Remove blocked items
                2. Calculate weights and shipping costs
                3. Show metrics summary
                4. Display processed data with highlights
                """,
                'highlight': 'metrics'
            },
            4: {
                'title': 'Export Results 💾',
                'content': """
                When you're satisfied with the results:
                
                1. Click 'Export to Excel'
                2. Download the processed file
                3. Save as CSV for final use
                
                Tip: Rows with missing weights are highlighted in red!
                """,
                'highlight': 'export'
            }
        }

    def render_tutorial(self):
        """Render the current tutorial step."""
        if not st.session_state.tutorial_active:
            return

        content = self.get_tutorial_content()
        current_step = st.session_state.tutorial_step
        max_steps = len(content) - 1

        # Create a container for the tutorial
        tutorial_container = st.sidebar.container()
        
        with tutorial_container:
            st.markdown("---")
            st.markdown("### 🎓 Tutorial Mode")
            
            # Display current step content
            step_content = content[current_step]
            st.markdown(f"#### {step_content['title']}")
            st.markdown(step_content['content'])
            
            # Navigation buttons
            cols = st.columns(3)
            
            with cols[0]:
                if current_step > 0:
                    if st.button("◀ Previous"):
                        st.session_state.tutorial_step -= 1
                        st.experimental_rerun()
                        
            with cols[1]:
                st.markdown(f"Step {current_step + 1}/{max_steps + 1}")
                
            with cols[2]:
                if current_step < max_steps:
                    if st.button("Next ▶"):
                        st.session_state.tutorial_step += 1
                        st.experimental_rerun()
                else:
                    if st.button("Finish 🎉"):
                        st.session_state.tutorial_active = False
                        st.experimental_rerun()

    def toggle_tutorial(self):
        """Toggle the tutorial state."""
        if st.sidebar.button(
            "Toggle Tutorial 🎓",
            help="Click to start/stop the interactive tutorial",
            type="primary" if not st.session_state.tutorial_active else "secondary"
        ):
            st.session_state.tutorial_active = not st.session_state.tutorial_active
            st.session_state.tutorial_step = 0
            st.experimental_rerun()
