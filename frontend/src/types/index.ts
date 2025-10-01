export interface Slide {
  id?: string;
  header: string;
  content: string;
  imageUrl?: string;
  image_url?: string; // Backend compatibility
}

export interface PresentationTemplate {
  id: string;
  name: string;
  description: string;
  preview: string;
  style: {
    backgroundColor: string;
    textColor: string;
    accentColor: string;
    fontFamily: string;
  };
}

export interface PresentationRequest {
  title: string;
  slides: number;
  template?: string;
  style?: 'professional' | 'creative' | 'minimal' | 'colorful';
  tone?: 'formal' | 'casual' | 'academic' | 'business';
  includeImages?: boolean;
  include_images?: boolean; // Backend compatibility
}

export interface PresentationResponse {
  slides: Slide[];
  title: string;
  presentation_id: string;
  generatedAt?: string;
  downloadUrl?: string;
  download_url?: string; // Backend compatibility
}

export interface GenerationProgress {
  step: 'generating' | 'fetching-images' | 'creating-slides' | 'complete';
  progress: number;
  message: string;
}

export interface ApiError {
  message: string;
  code?: string;
  details?: any;
}

export interface ApiResponse<T = any> {
  success: boolean;
  data?: T;
  message?: string;
  error?: string;
}
