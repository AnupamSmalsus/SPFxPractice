import { Web } from 'sp-pnp-js';

export const getJobs = async (props: any) => {
  const web = new Web(props.siteUrl);
  const jobsList = await web.lists.getById(props.JobListId).items.select('Id', 'Title', 'Description', 'Experience', 'Vacancy', 'Package', 'IsActive').getAll();
  return jobsList;
}

export const addJob = async (props: any, job: any) => {
  const web = new Web(props.siteUrl);
  const result = await web.lists.getById(props.JobListId).items.add(job);
  return result;
}

export const updateJob = async (props: any, jobId: number, job: any) => {
  const web = new Web(props.siteUrl);
  const result = await web.lists.getById(props.JobListId).items.getById(jobId).update(job);
  return result;
}

export const deleteJob = async (props: any, jobId: number) => {
  const web = new Web(props.siteUrl);
  const result = await web.lists.getById(props.JobListId).items.getById(jobId).recycle();
  return result;
}

export const addApplication = async (props: any, application: any) => {
  const web = new Web(props.siteUrl);
  const result = await web.lists.getById(props.ApplicationListId).items.add(application);
  return result;
}