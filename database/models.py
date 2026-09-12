from datetime import datetime, date
from sqlalchemy import String, Text, Integer, Date, DateTime, ForeignKey, Boolean, Float
from sqlalchemy.orm import Mapped, mapped_column, relationship
from .db import Base

class Class(Base):
    __tablename__ = "classes"
    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(100), unique=True)
    grade: Mapped[str | None] = mapped_column(String(50), nullable=True)
    study_direction: Mapped[str | None] = mapped_column(String(100), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    lessons = relationship("Lesson", back_populates="class_group")

class Lesson(Base):
    __tablename__ = "lessons"
    id: Mapped[int] = mapped_column(primary_key=True)
    title: Mapped[str] = mapped_column(String(200))
    subject: Mapped[str] = mapped_column(String(100), default="Nederlands")
    class_id: Mapped[int | None] = mapped_column(ForeignKey("classes.id"), nullable=True)
    date: Mapped[date | None] = mapped_column(Date, nullable=True)
    duration_minutes: Mapped[int] = mapped_column(Integer, default=50)
    topic: Mapped[str] = mapped_column(String(300), default="")
    description: Mapped[str] = mapped_column(Text, default="")
    status: Mapped[str] = mapped_column(String(50), default="draft")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    class_group = relationship("Class", back_populates="lessons")
    objectives = relationship("LearningObjective", back_populates="lesson", cascade="all, delete-orphan")
    phases = relationship("LessonPhase", back_populates="lesson", cascade="all, delete-orphan", order_by="LessonPhase.order_index")
    analyses = relationship("Analysis", back_populates="lesson", cascade="all, delete-orphan")
    retrievals = relationship("RetrievalSchedule", back_populates="lesson", cascade="all, delete-orphan")

class LearningObjective(Base):
    __tablename__ = "learning_objectives"
    id: Mapped[int] = mapped_column(primary_key=True)
    lesson_id: Mapped[int] = mapped_column(ForeignKey("lessons.id"))
    objective_text: Mapped[str] = mapped_column(Text)
    objective_type: Mapped[str] = mapped_column(String(50), default="knowledge")
    lesson = relationship("Lesson", back_populates="objectives")

class LessonPhase(Base):
    __tablename__ = "lesson_phases"
    id: Mapped[int] = mapped_column(primary_key=True)
    lesson_id: Mapped[int] = mapped_column(ForeignKey("lessons.id"))
    order_index: Mapped[int] = mapped_column(Integer, default=0)
    title: Mapped[str] = mapped_column(String(200))
    duration_minutes: Mapped[int] = mapped_column(Integer, default=5)
    teacher_activity: Mapped[str] = mapped_column(Text, default="")
    student_activity: Mapped[str] = mapped_column(Text, default="")
    activity_type: Mapped[str] = mapped_column(String(100), default="")
    materials: Mapped[str] = mapped_column(Text, default="")
    lesson = relationship("Lesson", back_populates="phases")

class Principle(Base):
    __tablename__ = "principles"
    id: Mapped[int] = mapped_column(primary_key=True)
    number: Mapped[int] = mapped_column(Integer, unique=True)
    name: Mapped[str] = mapped_column(String(250))
    short_description: Mapped[str] = mapped_column(Text)
    scores = relationship("PrincipleScore", back_populates="principle")

class Analysis(Base):
    __tablename__ = "analyses"
    id: Mapped[int] = mapped_column(primary_key=True)
    lesson_id: Mapped[int] = mapped_column(ForeignKey("lessons.id"))
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    overall_score: Mapped[float] = mapped_column(Float, default=0)
    max_score: Mapped[int] = mapped_column(Integer, default=48)
    summary: Mapped[str] = mapped_column(Text, default="")
    priority_1: Mapped[str] = mapped_column(Text, default="")
    priority_2: Mapped[str] = mapped_column(Text, default="")
    priority_3: Mapped[str] = mapped_column(Text, default="")
    lesson = relationship("Lesson", back_populates="analyses")
    principle_scores = relationship("PrincipleScore", back_populates="analysis", cascade="all, delete-orphan")

class PrincipleScore(Base):
    __tablename__ = "principle_scores"
    id: Mapped[int] = mapped_column(primary_key=True)
    analysis_id: Mapped[int] = mapped_column(ForeignKey("analyses.id"))
    principle_id: Mapped[int] = mapped_column(ForeignKey("principles.id"))
    score: Mapped[int] = mapped_column(Integer, default=0)
    status: Mapped[str] = mapped_column(String(20), default="red")
    evidence: Mapped[str] = mapped_column(Text, default="")
    strength: Mapped[str] = mapped_column(Text, default="")
    problem: Mapped[str] = mapped_column(Text, default="")
    recommendation: Mapped[str] = mapped_column(Text, default="")
    priority: Mapped[int] = mapped_column(Integer, default=0)
    analysis = relationship("Analysis", back_populates="principle_scores")
    principle = relationship("Principle", back_populates="scores")

class RetrievalSchedule(Base):
    __tablename__ = "retrieval_schedule"
    id: Mapped[int] = mapped_column(primary_key=True)
    lesson_id: Mapped[int] = mapped_column(ForeignKey("lessons.id"))
    topic: Mapped[str] = mapped_column(String(300))
    scheduled_date: Mapped[date] = mapped_column(Date)
    retrieval_type: Mapped[str] = mapped_column(String(100), default="quiz")
    question_count: Mapped[int] = mapped_column(Integer, default=3)
    completed: Mapped[bool] = mapped_column(Boolean, default=False)
    notes: Mapped[str] = mapped_column(Text, default="")
    lesson = relationship("Lesson", back_populates="retrievals")
