import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';

@Injectable({
  providedIn: 'root'
})
export class ExcelUploadService {
  private baseUrl = 'https://localhost:7250/api/ExcelUpload';

  constructor(private http: HttpClient) { }

  // Upload Excel file
  public uploadExcel(formData: FormData) {
    return this.http.post(`${this.baseUrl}`, formData);
  }

  // Get list of uploaded Excel file names (for dropdown)
  public getFileNames() {
    return this.http.get<string[]>(`${this.baseUrl}/files`);
  }

  // Get data from a specific Excel file
  public getExcelData(fileName: string) {
    return this.http.get<any[]>(`${this.baseUrl}/data?filename=${fileName}`);
  }
}
