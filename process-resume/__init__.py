import json, os, tempfile, traceback, logging
from pathlib import Path
import requests
import azure.functions as func

from .resume_processor import (
    extract_text_generic, call_model, sanitize_basename,
    export_profile_to_docx
)

def _bytes_to_docx(file_bytes: bytes, filename: str) -> tuple[str, bytes, bytes]:
    """
    Converts resume file bytes into extracted text, runs the model, and exports results to JSON and DOCX.
    Handles all intermediate steps and errors.
    Returns: (base filename, JSON bytes, DOCX bytes)
    """
    try:
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            in_path = td / filename
            in_path.write_bytes(file_bytes)

            # Extract text from the uploaded file
            try:
                text = extract_text_generic(in_path)
            except Exception as e:
                logging.error("Error extracting text: %s", e)
                raise RuntimeError("Failed to extract text from file") from e

            # Run the model on the extracted text
            try:
                profile = call_model(text)
            except Exception as e:
                logging.error("Error calling model: %s", e)
                raise RuntimeError("Failed to process resume with model") from e

            # Sanitize the base filename for output files
            base = sanitize_basename(profile.get("Name") or Path(filename).stem)

            json_path = td / f"{base}.json"
            # Serialize the profile to JSON and write to file
            try:
                json_bytes = json.dumps(profile, ensure_ascii=False, indent=2).encode("utf-8")
                json_path.write_bytes(json_bytes)
            except Exception as e:
                logging.error("Error serializing/writing JSON: %s", e)
                raise RuntimeError("Failed to serialize or write JSON") from e

            docx_path = td / f"{base}.docx"
            # Export the profile to DOCX format
            try:
                export_profile_to_docx(str(json_path), str(docx_path))
                docx_bytes = docx_path.read_bytes()
            except Exception as e:
                logging.error("Error exporting profile to DOCX: %s", e)
                raise RuntimeError("Failed to export profile to DOCX") from e

            return base, json_bytes, docx_bytes
    except Exception as e:
        logging.error("Error in _bytes_to_docx: %s", e)
        raise

from azure.storage.blob import BlobServiceClient
import os
 
def save_to_blob(file_bytes: bytes, blob_name: str):
    # Get connection string from Azure Function settings (Application Settings)
    connect_str = os.getenv("AzureWebJobsresumeconverter")
 
    # Create the BlobServiceClient
    blob_service_client = BlobServiceClient.from_connection_string(connect_str)
 
    # Get container client
    container_client = blob_service_client.get_container_client("resumeoutput")
 
    # Upload the file
    blob_client = container_client.get_blob_client(blob_name)
    blob_client.upload_blob(file_bytes, overwrite=True)
    return True

def main(blob: func.InputStream):
    """
    Azure Function Blob Trigger entry point.
    Triggered when a new blob is created in the configured container.
    Reads the blob bytes and processes the resume file.
    Uploads the processed DOCX file to SharePoint using an API.
    """
    try:
        # Read file bytes from the blob
        file_bytes = blob.read()
        filename = blob.name.split('/')[-1]  # Extract filename from blob path

        logging.info("Processing blob: %s (%d bytes)", filename, len(file_bytes))

        # Process the file and generate outputs
        try:
            base, json_bytes, docx_bytes = _bytes_to_docx(file_bytes, filename)
        except Exception as e:
            tb = traceback.format_exc()
            logging.error("Processing error: %s\n%s", e, tb)
            return

        # Upload the DOCX file to SharePoint
        docx_filename = f"{base}.docx"
        success = save_to_blob(docx_bytes, filename)
        if not success:
            logging.error("Failed to upload %s to SharePoint.", docx_filename)
            raise RuntimeError("SharePoint upload failed")
        else:
            logging.info("File %s uploaded to SharePoint successfully.", docx_filename)

    except Exception:
        tb = traceback.format_exc()
        logging.error("Unhandled error in blob trigger:\n%s", tb)