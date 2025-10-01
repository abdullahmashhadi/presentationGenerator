import { create } from 'zustand';
import { PresentationRequest, PresentationResponse, GenerationProgress, PresentationTemplate } from '@/types';

interface PresentationStore {
  // State
  currentRequest: Partial<PresentationRequest>;
  generationProgress: GenerationProgress | null;
  generatedPresentation: PresentationResponse | null;
  templates: PresentationTemplate[];
  isGenerating: boolean;
  error: string | null;
  
  // Actions
  setCurrentRequest: (request: Partial<PresentationRequest>) => void;
  setGenerationProgress: (progress: GenerationProgress | null) => void;
  setGeneratedPresentation: (presentation: PresentationResponse | null) => void;
  setTemplates: (templates: PresentationTemplate[]) => void;
  setIsGenerating: (isGenerating: boolean) => void;
  setError: (error: string | null) => void;
  resetState: () => void;
  updateRequestField: <K extends keyof PresentationRequest>(
    field: K,
    value: PresentationRequest[K]
  ) => void;
}

const initialRequest: Partial<PresentationRequest> = {
  title: '',
  slides: 5,
  template: 'modern',
  style: 'professional',
  tone: 'business',
  includeImages: true,
};

export const usePresentationStore = create<PresentationStore>((set, get) => ({
  // Initial state
  currentRequest: initialRequest,
  generationProgress: null,
  generatedPresentation: null,
  templates: [],
  isGenerating: false,
  error: null,

  // Actions
  setCurrentRequest: (request: Partial<PresentationRequest>) =>
    set({ currentRequest: request }),

  setGenerationProgress: (progress: GenerationProgress | null) =>
    set({ generationProgress: progress }),

  setGeneratedPresentation: (presentation: PresentationResponse | null) =>
    set({ generatedPresentation: presentation }),

  setTemplates: (templates: PresentationTemplate[]) =>
    set({ templates }),

  setIsGenerating: (isGenerating: boolean) =>
    set({ isGenerating }),

  setError: (error: string | null) =>
    set({ error }),

  resetState: () =>
    set({
      currentRequest: initialRequest,
      generationProgress: null,
      generatedPresentation: null,
      isGenerating: false,
      error: null,
    }),

  updateRequestField: <K extends keyof PresentationRequest>(field: K, value: PresentationRequest[K]) =>
    set((state) => ({
      currentRequest: {
        ...state.currentRequest,
        [field]: value,
      },
    })),
}));
