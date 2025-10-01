# AI Presentation Generator

A modern, professional AI-powered presentation generator built with Next.js and Flask.

## ✨ Features

### 🎯 Core Functionality

- **AI-Powered Content Generation** - Uses Google Gemini AI to create compelling slide content
- **Intelligent Image Integration** - Automatically fetches relevant images from Pexels API
- **Multiple Templates** - Choose from Modern Business, Creative Bold, Minimal Clean, and Academic styles
- **Customizable Styles** - Professional, Creative, Minimal, and Colorful presentation styles
- **Flexible Tones** - Formal, Casual, Academic, and Business communication styles

### 🚀 Advanced Features

- **Live Preview** - Preview your content before generating the full presentation
- **Progress Tracking** - Real-time progress updates during generation
- **Template Selection** - Visual template picker with style previews
- **Responsive Design** - Works perfectly on desktop, tablet, and mobile
- **Dark Mode Support** - Beautiful dark theme with automatic system detection
- **Error Handling** - Comprehensive error handling with user-friendly messages

### 🎨 Modern UI/UX

- **Glassmorphism Design** - Beautiful glass effects and gradients
- **Smooth Animations** - Framer Motion powered animations and transitions
- **Step-by-Step Wizard** - Intuitive 3-step presentation creation process
- **Professional Typography** - Inter font with perfect spacing and hierarchy
- **Accessibility** - Built with accessibility best practices

## 🏗️ Architecture

### Frontend (Next.js 14)

- **Framework**: Next.js 14 with App Router
- **Language**: TypeScript for type safety
- **Styling**: Tailwind CSS with custom design system
- **Animations**: Framer Motion for smooth transitions
- **State Management**: Zustand for lightweight state management
- **API Client**: Axios with error handling and retries
- **Icons**: Lucide React for consistent iconography

### Backend (Flask)

- **Framework**: Flask with modern Python patterns
- **AI Integration**: Google Gemini 1.5 Flash for content generation
- **Image API**: Pexels API for relevant stock photos
- **Document Generation**: python-pptx for PowerPoint creation
- **CORS**: Flask-CORS for secure cross-origin requests
- **Data Validation**: Dataclasses and type hints
- **Error Handling**: Comprehensive error responses

## 🚀 Quick Start

### Prerequisites

- Node.js 18+ and npm
- Python 3.9+
- Google Gemini API key
- Pexels API key (optional, for images)

### Backend Setup

1. **Clone and navigate to backend**:

   ```bash
   cd backend
   ```

2. **Create virtual environment**:

   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**:

   ```bash
   pip install -r requirements.txt
   ```

4. **Configure environment variables**:

   ```bash
   # Create .env file in backend directory
   API_KEY=your_gemini_api_key_here
   PEXELS_API_KEY=your_pexels_api_key_here  # Optional
   ```

5. **Run the backend**:
   ```bash
   python app.py
   ```
   Backend will be available at `http://localhost:5000`

### Frontend Setup

1. **Navigate to frontend**:

   ```bash
   cd frontend
   ```

2. **Install dependencies**:

   ```bash
   npm install
   ```

3. **Configure environment**:

   ```bash
   # Create .env.local file in frontend directory
   NEXT_PUBLIC_BACKEND_URL=http://localhost:5000
   ```

4. **Run the frontend**:
   ```bash
   npm run dev
   ```
   Frontend will be available at `http://localhost:3000`

## 📝 API Documentation

### Backend Endpoints

#### Health Check

```
GET /health
```

Returns service health status and version.

#### Get Templates

```
GET /templates
```

Returns available presentation templates.

#### Preview Content

```
POST /preview
Content-Type: application/json

{
  "title": "Your Presentation Title",
  "slides": 5
}
```

Generates slide content preview without creating presentation.

#### Generate Presentation

```
POST /generate
Content-Type: application/json

{
  "title": "Your Presentation Title",
  "slides": 5,
  "template": "modern",
  "style": "professional",
  "tone": "business",
  "includeImages": true
}
```

Generates complete presentation with PowerPoint file.

#### Download Presentation

```
GET /download/{presentation_id}
```

Downloads the generated PowerPoint file.

## 🎨 Customization

### Adding New Templates

1. **Define template in backend**:

   ```python
   PresentationTemplate(
       id="custom",
       name="Custom Template",
       description="Your custom description",
       preview="/templates/custom.png",
       style={
           "backgroundColor": "#ffffff",
           "textColor": "#2d3748",
           "accentColor": "#your-color",
           "fontFamily": "Your Font"
       }
   )
   ```

2. **Add template logic** in `create_presentation` method.

### Modifying Styles

Update the Tailwind configuration in `frontend/tailwind.config.js` to customize colors, fonts, and animations.

### Extending AI Prompts

Modify the prompt generation in `backend/app.py` to customize how Gemini AI generates content for different styles and tones.

## 🚀 Deployment

### Backend Deployment (Vercel)

1. **Deploy to Vercel**:

   ```bash
   cd backend
   vercel --prod
   ```

2. **Set environment variables** in Vercel dashboard:
   - `API_KEY`: Your Gemini API key
   - `PEXELS_API_KEY`: Your Pexels API key

### Frontend Deployment (Vercel)

1. **Update backend URL**:

   ```bash
   # In frontend/.env.local
   NEXT_PUBLIC_BACKEND_URL=https://your-backend-url.vercel.app
   ```

2. **Deploy to Vercel**:
   ```bash
   cd frontend
   vercel --prod
   ```

### Alternative Deployment Options

- **Railway**: Great for both frontend and backend
- **Heroku**: Traditional platform with easy deployment
- **DigitalOcean App Platform**: Containerized deployment
- **AWS/GCP**: Full cloud platform deployment

## 🔧 Environment Variables

### Backend (.env)

```env
API_KEY=your_gemini_api_key
PEXELS_API_KEY=your_pexels_api_key  # Optional for images
FLASK_ENV=production  # For production
```

### Frontend (.env.local)

```env
NEXT_PUBLIC_BACKEND_URL=http://localhost:5000  # Development
# NEXT_PUBLIC_BACKEND_URL=https://your-backend.vercel.app  # Production
```

## 🤝 Contributing

1. Fork the repository
2. Create your feature branch: `git checkout -b feature/amazing-feature`
3. Commit your changes: `git commit -m 'Add amazing feature'`
4. Push to the branch: `git push origin feature/amazing-feature`
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **Google Gemini AI** for powerful content generation
- **Pexels** for beautiful stock photography
- **Next.js Team** for the amazing React framework
- **Flask Community** for the lightweight Python framework
- **Tailwind CSS** for the utility-first CSS framework
- **Framer Motion** for smooth animations

## 📞 Support

For support, email support@presentationgenerator.com or create an issue in this repository.

---

Built with ❤️ by the Presentation Generator Team
