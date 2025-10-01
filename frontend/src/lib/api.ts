import axios from 'axios';
import { PresentationRequest, PresentationResponse, Slide, ApiResponse } from '@/types';

const API_BASE_URL = process.env.NEXT_PUBLIC_BACKEND_URL || '/api';

const api = axios.create({
  baseURL: API_BASE_URL,
  timeout: 60000,
  headers: {
    'Content-Type': 'application/json',
  },
});

export class ApiService {
  static async generatePresentation(data: PresentationRequest): Promise<ApiResponse<PresentationResponse>> {
    try {
      // Transform the data to match backend expectations
      const requestData = {
        title: data.title,
        slides: data.slides,
        template: data.template || 'modern',
        style: data.style || 'professional',
        tone: data.tone || 'business',
        includeImages: data.includeImages ?? true,
      };
      
      const response = await api.post('/generate', requestData);
      return response.data;
    } catch (error) {
      throw this.handleError(error);
    }
  }

  static async downloadPresentation(presentationId: string): Promise<Blob> {
    try {
      const response = await api.get(`/download/${presentationId}`, {
        responseType: 'blob',
      });
      return response.data;
    } catch (error) {
      throw this.handleError(error);
    }
  }

  static async getTemplates(): Promise<ApiResponse<any[]>> {
    try {
      const response = await api.get('/templates');
      return response.data;
    } catch (error) {
      throw this.handleError(error);
    }
  }

  static async previewSlides(data: Pick<PresentationRequest, 'title' | 'slides'>): Promise<ApiResponse<{ slides: Slide[] }>> {
    try {
      const response = await api.post('/preview', data);
      return response.data;
    } catch (error) {
      throw this.handleError(error);
    }
  }

  static async healthCheck(): Promise<ApiResponse<{ status: string; version: string }>> {
    try {
      const response = await api.get('/health');
      return response.data;
    } catch (error) {
      throw this.handleError(error);
    }
  }

  private static handleError(error: any): Error {
    if (error.response) {
      // Server responded with error status
      const message = error.response.data?.error || error.response.data?.message || 'Server error occurred';
      return new Error(`${message} (${error.response.status})`);
    } else if (error.request) {
      // Request made but no response
      return new Error('Network error - please check your connection');
    } else {
      // Something else happened
      return new Error(error.message || 'An unexpected error occurred');
    }
  }
}

export default ApiService;
