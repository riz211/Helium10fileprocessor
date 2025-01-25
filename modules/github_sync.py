"""GitHub synchronization module for the application."""
import os
import streamlit as st
from git import Repo
from git.exc import GitCommandError

class GitHubSync:
    """Manages GitHub repository synchronization."""

    def __init__(self, repo_url: str = "https://github.com/riz211/Helium10fileprocessor.git"):
        """Initialize the GitHub sync manager."""
        self.repo_url = repo_url
        self.token = os.environ.get('GITHUB_TOKEN')
        self.repo = None
        try:
            self.repo = Repo('.')
        except:
            st.error("Unable to initialize Git repository")

    def sync_changes(self) -> tuple[bool, str]:
        """
        Synchronize local changes with GitHub repository.

        Returns:
            tuple: (success status, message)
        """
        if not self.token:
            return False, "GitHub token not found. Please check your environment variables."

        if not self.repo:
            return False, "Git repository not initialized."

        try:
            # Configure remote with token
            remote_url = f"https://{self.token}@github.com/riz211/Helium10fileprocessor.git"

            # Update remote URL with token
            if 'origin' in [remote.name for remote in self.repo.remotes]:
                self.repo.delete_remote('origin')
            self.repo.create_remote('origin', remote_url)

            # Explicitly add requirements.txt if it exists
            requirements_path = "requirements.txt"
            if os.path.exists(requirements_path):
                self.repo.index.add([requirements_path])

            # Stage all other changes
            self.repo.git.add(A=True)

            # Commit if there are changes
            if self.repo.is_dirty() or len(self.repo.untracked_files) > 0:
                self.repo.index.commit("Update application files including requirements.txt")

            # Make sure we're on the main branch
            if 'main' not in self.repo.heads:
                self.repo.create_head('main')
            self.repo.heads.main.checkout()

            # Push changes
            origin = self.repo.remote('origin')
            origin.push('main:main')

            return True, "Successfully synchronized changes with GitHub!"

        except GitCommandError as e:
            return False, f"Git error: {str(e)}"
        except Exception as e:
            return False, f"Error syncing with GitHub: {str(e)}"

    def render_sync_button(self):
        """Render the sync button in the Streamlit interface."""
        st.sidebar.markdown("---")
        st.sidebar.subheader("📤 GitHub Sync")

        if st.sidebar.button("Sync with GitHub", help="Push current changes to GitHub repository"):
            with st.spinner("Syncing changes..."):
                success, message = self.sync_changes()

                if success:
                    st.sidebar.success(message)
                else:
                    st.sidebar.error(message)