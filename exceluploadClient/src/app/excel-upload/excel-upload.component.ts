import { Component, OnInit } from '@angular/core'; // adjust path if needed
import { ExcelUploadService } from '../excel-upload.service';
import { HttpClient } from '@angular/common/http';

@Component({
  selector: 'app-excel-upload',
  templateUrl: './excel-upload.component.html'
})
export class ExcelUploadComponent implements OnInit {

  selectedFile!: File;
  uploadSuccess: boolean | null = null;
  errorMessage: string = '';

  uploadedFiles: string[] = [];
  selectedFileName: string = '';
  excelData: any[] = [];
  constructor(private excelService: ExcelUploadService, private http: HttpClient) { }

  ngOnInit(): void {
    this.fetchUploadedFiles();
  }

  onFileSelected(event: any): void {
    this.selectedFile = event.target.files[0];
  }

  onUpload(): void {
    if (!this.selectedFile) {
      this.errorMessage = 'Please select a file.';
      this.uploadSuccess = false;
      return;
    }

    const formData = new FormData();
    formData.append('file', this.selectedFile);

    this.excelService.uploadExcel(formData).subscribe({
      next: () => {
        this.uploadSuccess = true;
        this.errorMessage = '';
        this.fetchUploadedFiles();
      },
      error: (err) => {
        this.uploadSuccess = false;
        this.errorMessage = err.error || 'Upload failed.';
      }
    });
  }

  fetchUploadedFiles(): void {
    this.excelService.getFileNames().subscribe({
      next: (files) => {
        this.uploadedFiles = files;
      },
      error: (err) => {
        console.error('Error fetching uploaded files:', err);
      }
    });
  }

  onFileChange(): void {
    if (!this.selectedFile) {
      this.excelData = [];
      return;
    }
  
    this.http.get<any[]>(`https://localhost:7250/api/ExcelUpload/data?filename=${this.selectedFile}`)
      .subscribe({
        next: (data) => {
          this.excelData = data;
        },
        error: (error) => {
          console.error('Error fetching data:', error);
          this.excelData = [];
        }
      });
  }
  

  onFileSelectChange(): void {
    if (!this.selectedFileName) return;

    this.excelService.getExcelData(this.selectedFileName).subscribe({
      next: (data) => {
        this.excelData = data;
      },
      error: (err) => {
        console.error('Error loading file data:', err);
        this.errorMessage = 'Failed to load data.';
        this.uploadSuccess = false;
      }
    });
  }
}
