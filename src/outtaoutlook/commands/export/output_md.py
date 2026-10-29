"""Export messages as Markdown files."""

import logging
import unicodedata
from collections.abc import Sequence
from decimal import Decimal
from pathlib import Path
from typing import Self
from urllib.parse import unquote

import yaml
from html_to_markdown import ConversionOptions
from html_to_markdown import convert as convert_to_markdown

from outtaoutlook.graphapi import GraphMessageT, log_message_notice

from .output_core import ExportStat, OutputWorkerCore
from .shared import DEFAULT_SELECT_MESSAGE_FIELDS

######################################################################
# Constants


# Message fields we require for exporting.
SELECT_MESSAGE_FIELDS = ",".join(
    sorted(
        (
            *DEFAULT_SELECT_MESSAGE_FIELDS,
            "body",
        ),
    ),
)


######################################################################
# Types and classes


FrontmatterT = dict[str, str | int | Decimal | bool | Sequence[str | int]]


######################################################################
# Public module class


class OutputWorkerMarkdown(OutputWorkerCore):
    """Implements the `OutputWorkerCore` for Markdown output."""

    ######################################################################
    # Internal methods

    @staticmethod
    def _clean_tags(tags: list[str]) -> list[str]:
        """Normalize tags to follow rules Obsidian defines.

        https://obsidian.md/help/tags

        Args:
            tags (list[str]): List of tags.

        Returns:
            list[str]: Updated list of tags.
        """
        for i, _tag in enumerate(tags):
            tag = _tag.strip()

            if " " in tag:
                tag = tag.title().replace(" ", "")

            tag = "".join(
                # Thanks Gemini
                char
                for char in tag
                if unicodedata.category(char)[0] in "LNMS" or char in "_-/\u200d"
            )

            if tag.isdecimal():
                tag = f"_{tag}"

            tags[i] = tag

        return tags

    def _init_file_frontmatter(self, message: GraphMessageT) -> FrontmatterT:
        """Initialize file frontmatter with expanded metadata properties.

        Args:
            message (GraphMessageT): A message object.

        Returns:
            FrontmatterT: Initialized frontmatter dictionary.
        """

        def is_float(value: str) -> bool:
            try:
                float(value)
            except ValueError:
                return False
            return True

        frontmatter: FrontmatterT = {}

        metaprop = self._expand_vars_dict(
            self.export_opts.meta_props or {},
            self.get_message_metaprop_vars(message),
        )

        # Properties support some scalar types
        for key, value in metaprop.items():
            if not isinstance(value, str):
                frontmatter[key] = value
            elif value.isdecimal():
                frontmatter[key] = int(value)
            elif "." in value and is_float(value):
                frontmatter[key] = Decimal(value)
            elif value.lower() in ("true", "false"):
                frontmatter[key] = value.lower() == "true"
            else:
                frontmatter[key] = value

        return frontmatter

    def _append_md_tags(
        self,
        message: GraphMessageT,
        file_content: str,
        frontmatter: FrontmatterT,
    ) -> str:
        """Append tags to file content and/or frontmatter.

        Args:
            message (GraphMessageT): A message object.
            file_content (str): File content.
            frontmatter (FrontmatterT): Frontmatter dictionary. Frontmatter tags will be added here if applicable.

        Returns:
            str
        """
        tag_vars = self.get_message_tag_vars(message)
        body_tags = self._expand_vars_list(self.export_opts.md_body_tags or [], tag_vars)
        frontmatter_tags = self._expand_vars_list(self.export_opts.md_tags or [], tag_vars)

        if not self.export_opts.md_place_tags_frontmatter:
            # Place all tags in body
            body_tags.extend(frontmatter_tags)
            frontmatter_tags = []

        # Format as valid tags
        body_tags = self._clean_tags(body_tags)
        frontmatter_tags = self._clean_tags(frontmatter_tags)

        # Remove duplicates while keeping order
        body_tags = list(dict.fromkeys(body_tags))
        frontmatter_tags = list(dict.fromkeys(frontmatter_tags))

        if frontmatter_tags:
            frontmatter["tags"] = frontmatter_tags  # modifies caller

        if body_tags:
            tag_block = " ".join(f"#{tag}" for tag in body_tags)

            if self.export_opts.md_place_tags_body_end:
                file_content += "\n\n" + tag_block
            else:
                file_content = tag_block + "\n\n" + file_content

        return file_content

    @staticmethod
    def _convert_to_markdown(body_content: str) -> str | None:
        """Convert message content to Markdown.

        Args:
            body_content (str): The HTML body content of the message.

        Returns:
            str|None: File content or None if message has no body or Markdown conversion fails.
        """
        # https://github.com/matthewwithanm/python-markdownify
        # from markdownify import markdownify as html_to_md
        # file_content = html_to_md(body_content)

        # https://github.com/xberg-io/html-to-markdown
        # https://pypi.org/project/html-to-markdown/
        # https://docs.html-to-markdown.xberg.io/reference/api-python/

        # Disable metadata extraction to prevent Microsoft Word custom <meta> tags from being added to the frontmatter.
        # I couldn't find a way to parse it and suppress the frontmatter altogether :/
        # Everything is escaped presuming emails aren't meant to include Markdown entities, such as hashtags.
        opts = ConversionOptions(
            extract_metadata=False,
            escape_ascii=True,
            escape_asterisks=True,
            escape_misc=True,
            escape_underscores=True,
        )

        convert_result = convert_to_markdown(body_content, opts)

        return convert_result.content

    @staticmethod
    def _attachment_abs_to_rel(abs_paths: list[Path], file_dir: Path, root_out_dir: Path) -> list[str]:
        """Convert absolute paths into paths relative to the file directory or the root output directory.

        Args:
            abs_paths (list[Path]): List of absolute paths to convert.
            file_dir (Path): Directory of the output file.
            root_out_dir (Path): Root output directory.

        Returns:
            list[str]: Paths relative to root output directory or URIs (file://) for absolute paths.
        """
        always_rel = False
        rel_paths: list[str] = []

        # Assuming we are exporting Markdown files into an Obsidian vault...
        # For links without a path, Obsidian seems to search:
        #
        # 1. Note folder
        # 2. Note subfolders
        # 3. Vault root
        # 4. Vault subfolders, alphabetically
        #
        # Outside the note directory and vault, use an absolute file URI.

        for abs_path in abs_paths:
            assert abs_path.is_absolute()  # noqa: S101 ; as_uri fails otherwise
            if always_rel:
                try:
                    rel_path = str(abs_path.relative_to(root_out_dir, walk_up=True))
                except ValueError:
                    rel_path = abs_path.as_uri()
            elif abs_path.is_relative_to(file_dir) or abs_path.is_relative_to(root_out_dir):
                rel_path = abs_path.name
            else:
                rel_path = abs_path.as_uri()

            rel_paths.append(rel_path)

        return rel_paths

    @staticmethod
    def _md_link_text(rel_path_or_uri: str, *, inline: bool) -> str:
        """Generate a Markdown link for a relative path or URI.

        Args:
            rel_path_or_uri (str): Relative path or URI to link.
            inline (bool): Whether the link is inline or not.

        Returns:
            str: Link notation.
        """
        if rel_path_or_uri.startswith("file://"):
            # Links external to the root output directory are normal markdown links
            abs_path = Path(rel_path_or_uri[7:])
            return f"[{unquote(abs_path.name)}]({rel_path_or_uri})"

        return f"![[{rel_path_or_uri}]]" if inline else f"[[{rel_path_or_uri}]]"

    ######################################################################
    # Overridden methods

    def update_messages_params(self: Self, params: dict[str, str | int]) -> None:
        """Update parameters used in the '/messages' Graph API request.

        Args:
            params (dict[str, str | int]): Dictionary of parameters to update.
        """
        params["$select"] = SELECT_MESSAGE_FIELDS

    def get_file_path(
        self: Self,
        message: GraphMessageT,
        suffix: str | None = None,  # noqa: ARG002
    ) -> Path:
        """Generate a prospective path to an output file.

        Args:
            message (GraphMessageT): The message.
            suffix (str | None): Optional file suffix (e.g., ".md", ".eml").

        Returns:
            Path
        """
        return super().get_file_path(message, ".md")

    def export_write_message(  # noqa: C901, PLR0911, PLR0912
        self: Self,
        message: GraphMessageT,
        file_path: Path,
        attachment_paths: list[Path] | None,
    ) -> int:
        """Export a message as a Markdown file.

        Args:
            message (GraphMessageT): The message.
            file_path (Path): Prospective path to the output file to write.
            attachment_paths (list[Path] | None): List of attachment file paths or None if no attachments.

        Returns:
            int: The number of messages processed. Typically 1, but can be 0 if queued for later and >0 when processed.

        Raises:
            requests.RequestException derived error (e.g. ConnectionError, JSONDecodeError, HTTPError, etc.)
            OSError derived error (e.g. file system errors, permission errors, etc.)
        """

        def _skip_message(skip_reason: ExportStat, level: int) -> int:
            self.export_stats.incr(skip_reason)
            log_message_notice(message, skip_reason.value, level)
            return 1

        file_date = message["messageDateTime"]
        body_content: str | None = message.get("body", {}).get("content")
        content_type: str | None = message.get("body", {}).get("contentType")
        has_attachments: bool | None = message.get("hasAttachments")
        file_body = body_content

        if body_content:
            if content_type == "html":
                file_body = self._convert_to_markdown(body_content)
            elif content_type != "text":
                log_message_notice(message, f"field('contentType'): unexpected type: '{content_type}'")

        if not file_body:
            if body_content is None:
                return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_BODY, logging.WARNING)
            if file_body is None:
                return _skip_message(ExportStat.STAT_SKIPPED_MARKDOWN_CONVERSION_FAILED, logging.WARNING)
            if has_attachments and self.export_opts.with_attachments:
                file_body = ""  # content will be links to attachments
            if not has_attachments:
                return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_NO_ATTACHMENTS, logging.INFO)
            if not self.export_opts.with_attachments:
                return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_JUST_ATTACHMENTS, logging.INFO)

        if file_body is None:
            return _skip_message(ExportStat.STAT_SKIPPED_NO_MESSAGE_CONTENT_UNKNOWN, logging.ERROR)

        file_content: list[str] = []
        attachment_links: list[str] = []

        # Link to attachments
        if attachment_paths:
            inline = len(attachment_paths) == 1 and not file_body  # File consists of a single inline link
            rel_links = self._attachment_abs_to_rel(attachment_paths, file_path.parent, self.export_opts.root_out_dir)
            attachment_links = [self._md_link_text(rel_path_or_uri, inline=inline) for rel_path_or_uri in rel_links]

        frontmatter = self._init_file_frontmatter(message)

        file_body = self._append_md_tags(message, file_body, frontmatter)

        if frontmatter:
            # Frontmatter goes first
            file_content = [
                "---",
                yaml.safe_dump(frontmatter, sort_keys=False).strip(),
                "---",
                *file_content,
            ]

        if attachment_links:
            file_content.extend(attachment_links)
            file_content.append("")  # blank line

        # File body goes last
        file_content.append(file_body)

        file_text = "\n".join(file_content).strip() + "\n"

        if self.export_opts.prime_cache:
            return 1  # only prime the cache, do not write files

        self.write_output_file(file_path, file_text, file_date)

        self.export_stats.incr(ExportStat.STAT_MESSAGE_EXPORTED)

        return 1
